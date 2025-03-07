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

#include "docpropscore.h"

#include <QDateTime>
#include <QDebug>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_YXLSX

inline const QLatin1String kCoreProperties("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
inline const QLatin1String kPLElements("http://purl.org/dc/elements/1.1/");
inline const QLatin1String kPLTerms("http://purl.org/dc/terms/");
inline const QLatin1String kPLDcmiType("http://purl.org/dc/dcmitype/");
inline const QLatin1String kW3SchemaInstance("http://www.w3.org/2001/XMLSchema-instance");

const QHash<QString, QString> DocPropsCore::kElementNamespaceHash { { QStringLiteral("creator"), kPLElements },
    { QStringLiteral("lastModifiedBy"), kCoreProperties }, { QStringLiteral("created"), kPLTerms }, { QStringLiteral("modified"), kPLTerms } };

DocPropsCore::DocPropsCore(OperationMode mode)
    : AbstractOOXmlFile { mode }
{
}

bool DocPropsCore::SetProperty(const QString& name, const QString& value)
{
    static const QSet<QString> valid_keys { QStringLiteral("creator"), QStringLiteral("lastModifiedBy"), QStringLiteral("created"),
        QStringLiteral("modified") };

    if (!valid_keys.contains(name)) {
        return false;
    }

    if (value.isEmpty()) {
        return false;
    }

    property_hash_[name] = value;
    return true;
}

void DocPropsCore::ComposeXml(QIODevice* device) const
{
    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QLatin1String("1.0"), true);
    writer.writeStartElement(QLatin1String("cp:coreProperties"));

    writer.writeNamespace(kCoreProperties, QLatin1String("cp"));
    writer.writeNamespace(kPLElements, QLatin1String("dc"));
    writer.writeNamespace(kPLTerms, QLatin1String("dcterms"));
    writer.writeNamespace(kPLDcmiType, QLatin1String("dcmitype"));
    writer.writeNamespace(kW3SchemaInstance, QLatin1String("xsi"));

    auto it = property_hash_.constFind(QStringLiteral("creator"));
    writer.writeTextElement(kPLElements, QLatin1String("creator"), it != property_hash_.constEnd() ? it.value() : QLatin1String("YXlsx Library"));

    it = property_hash_.constFind(QStringLiteral("lastModifiedBy"));
    writer.writeTextElement(kCoreProperties, QLatin1String("lastModifiedBy"), it != property_hash_.constEnd() ? it.value() : QLatin1String("YXlsx Library"));

    auto WriteTimeElement = [&](QLatin1String name, const QString& value) {
        writer.writeStartElement(kPLTerms, name);
        writer.writeAttribute(kW3SchemaInstance, QLatin1String("type"), QLatin1String("dcterms:W3CDTF"));
        writer.writeCharacters(value);
        writer.writeEndElement();
    };

    const auto created { property_hash_.constFind(QStringLiteral("created")) };
    WriteTimeElement(QLatin1String("created"), created != property_hash_.constEnd() ? created.value() : QDateTime::currentDateTime().toString(Qt::ISODate));
    WriteTimeElement(QLatin1String("modified"), QDateTime::currentDateTime().toString(Qt::ISODate));

    writer.writeEndElement(); // cp:coreProperties
    writer.writeEndDocument();
}

bool DocPropsCore::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);

    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        const auto namespace_uri { reader.namespaceUri().toString() };
        const auto name { reader.name().toString() };

        auto it = kElementNamespaceHash.constFind(name);
        if (it != kElementNamespaceHash.constEnd() && it.value() == namespace_uri) {
            SetProperty(name, reader.readElementText());
        }
    }

    if (reader.hasError()) {
        qDebug() << "DocPropsCore Error reading doc props app file:" << reader.errorString();
        return false;
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
