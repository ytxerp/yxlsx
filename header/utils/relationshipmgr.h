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

#ifndef YXLSX_RELATIONSHIPMGR_H
#define YXLSX_RELATIONSHIPMGR_H

#include <QtCore/qxmlstream.h>

#include <QHash>
#include <QIODevice>

#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

struct Relationship {
    QString id {};
    QString type {};
    QString target {};
    QString target_mode {};
};

class RelationshipMgr {
public:
    RelationshipMgr();

    QList<Relationship> GetDocumentRelationship(const QString& relative_type) const;
    QList<Relationship> GetPackageRelationship(const QString& relative_type) const;
    QList<Relationship> GetMSPackageRelationship(const QString& relative_type) const;
    QList<Relationship> GetWorksheetRelationship(const QString& relative_type) const;

    void SetDocumentRelationship(const QString& relative_type, const QString& target);
    void SetPackageRelationship(const QString& relative_type, const QString& target);
    void SetMSPackageRelationship(const QString& relative_type, const QString& target);
    void SetWorksheetRelationship(const QString& relative_type, const QString& target, const QString& target_mode = QString());

    Relationship GetRelationshipByID(const QString& id) const;

    bool ReadByteArray(const QByteArray& data);
    QByteArray WriteByteArray() const;

    inline void Clear() { relationship_hash_.clear(); }
    inline int Count() const { return relationship_hash_.count(); }
    inline bool IsEmpty() const { return relationship_hash_.isEmpty(); }

private:
    QList<Relationship> GetRelationshipByType(const QString& type) const;
    void SetRelationship(const QString& type, const QString& target, const QString& target_mode = QString());

    void ComposeXml(QIODevice* device) const;
    bool ParseXml(QIODevice* device);
    bool ParseRelationship(const QXmlStreamAttributes& attrs);
    void ComposeRelationship(QXmlStreamWriter& writer, const Relationship& relation) const;

private:
    QHash<QString, Relationship> relationship_hash_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_RELATIONSHIPMGR_H
