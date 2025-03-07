/*
 * MIT License
 *
 * Copyright (c) 2024 YTX
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#include "relationshipmgr.h"

#include <QBuffer>
#include <QDebug>

QT_BEGIN_NAMESPACE_YXLSX

inline const QString kOxDocument { QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships") };
inline const QString kMsOffice { QStringLiteral("http://schemas.microsoft.com/office/2006/relationships") };
inline const QString kOxPackage { QStringLiteral("http://schemas.openxmlformats.org/package/2006/relationships") };

RelationshipMgr::RelationshipMgr() { }

QList<Relationship> RelationshipMgr::GetDocumentRelationship(const QString& relative_type) const { return GetRelationshipByType(kOxDocument + relative_type); }

void RelationshipMgr::SetDocumentRelationship(const QString& relative_type, const QString& target) { SetRelationship(kOxDocument + relative_type, target); }

QList<Relationship> RelationshipMgr::GetMSPackageRelationship(const QString& relative_type) const { return GetRelationshipByType(kMsOffice + relative_type); }

void RelationshipMgr::SetMSPackageRelationship(const QString& relative_type, const QString& target) { SetRelationship(kMsOffice + relative_type, target); }

QList<Relationship> RelationshipMgr::GetPackageRelationship(const QString& relative_type) const { return GetRelationshipByType(kOxPackage + relative_type); }

void RelationshipMgr::SetPackageRelationship(const QString& relative_type, const QString& target) { SetRelationship(kOxPackage + relative_type, target); }

QList<Relationship> RelationshipMgr::GetWorksheetRelationship(const QString& relative_type) const { return GetRelationshipByType(kOxDocument + relative_type); }

void RelationshipMgr::SetWorksheetRelationship(const QString& relative_type, const QString& target, const QString& target_mode)
{
    SetRelationship(kOxDocument + relative_type, target, target_mode);
}

QList<Relationship> RelationshipMgr::GetRelationshipByType(const QString& type) const
{
    QList<Relationship> res {};
    for (const Relationship& ship : relationship_hash_) {
        if (ship.type == type)
            res.append(ship);
    }
    return res;
}

void RelationshipMgr::SetRelationship(const QString& type, const QString& target, const QString& target_mode)
{
    Relationship relation {};
    relation.id = QStringLiteral("rId%1").arg(relationship_hash_.size() + 1);
    relation.type = type;
    relation.target = target;
    relation.target_mode = target_mode;

    relationship_hash_.insert(relation.id, relation);
}

void RelationshipMgr::ComposeXml(QIODevice* device) const
{
    if (!device->isWritable()) {
        qWarning() << "Device is not writable.";
        return;
    }

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QLatin1String("1.0"), true);
    writer.writeStartElement(QLatin1String("Relationships"));
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/package/2006/relationships"));

    for (const Relationship& relationship : relationship_hash_) {
        ComposeRelationship(writer, relationship);
    }

    writer.writeEndElement(); // Relationships
    writer.writeEndDocument();
}

void RelationshipMgr::ComposeRelationship(QXmlStreamWriter& writer, const Relationship& relation) const
{
    writer.writeStartElement(QLatin1String("Relationship"));
    writer.writeAttribute(QLatin1String("Id"), relation.id);
    writer.writeAttribute(QLatin1String("Type"), relation.type);
    writer.writeAttribute(QLatin1String("Target"), relation.target);

    // Optional attribute
    if (!relation.target_mode.isEmpty()) {
        writer.writeAttribute(QLatin1String("TargetMode"), relation.target_mode);
    }

    writer.writeEndElement(); // Relationship
}

QByteArray RelationshipMgr::WriteByteArray() const
{
    QByteArray data {};
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    ComposeXml(&buffer);

    return data;
}

bool RelationshipMgr::ParseXml(QIODevice* device)
{
    Clear();
    QXmlStreamReader reader(device);
    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        if (reader.name() == QStringLiteral("Relationship")) {
            const auto attrs { reader.attributes() };
            if (!ParseRelationship(attrs)) {
                return false; // Parsing failed, exit early
            }
        }
    }

    if (reader.hasError()) {
        qWarning() << "XML Parsing Error:" << reader.errorString();
        return false;
    }

    return true;
}

bool RelationshipMgr::ParseRelationship(const QXmlStreamAttributes& attrs)
{
    // Extract attributes safely
    const QString id { attrs.value(QLatin1String("Id")).toString() };
    const QString type { attrs.value(QLatin1String("Type")).toString() };
    const QString target { attrs.value(QLatin1String("Target")).toString() };

    // Check required attributes
    if (id.isEmpty() || type.isEmpty() || target.isEmpty()) {
        qWarning() << "Missing required attributes in <Relationship>";
        return false;
    }

    // Optional attribute
    const QString target_mode = attrs.value(QLatin1String("TargetMode")).toString();

    // Store relationship
    Relationship relationship { id, type, target, target_mode };
    relationship_hash_.insert(id, relationship);
    return true;
}

bool RelationshipMgr::ReadByteArray(const QByteArray& data)
{
    QBuffer buffer {};
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);
    return ParseXml(&buffer);
}

Relationship RelationshipMgr::GetRelationshipByID(const QString& id) const { return relationship_hash_.value(id, Relationship()); }

QT_END_NAMESPACE_YXLSX
