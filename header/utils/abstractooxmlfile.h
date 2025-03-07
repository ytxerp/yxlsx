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

/*
 * Note:
 * If mainly using QStringView, QStringLiteral is more appropriate.
 * If using QAnyStringView and processing pure ASCII text, QLatin1String might be more lightweight and efficient.
 */

#ifndef YXLSX_ABSTRACTOOXMLFILE_H
#define YXLSX_ABSTRACTOOXMLFILE_H

#include <QSharedPointer>

#include "namespace.h"
#include "relationshipmgr.h"

QT_BEGIN_NAMESPACE_YXLSX

enum class OperationMode { kCreateNew, kLoadExisting };

class AbstractOOXmlFile {
public:
    virtual ~AbstractOOXmlFile();

    virtual void ComposeXml(QIODevice* device) const = 0;
    virtual bool ParseXml(QIODevice* device) = 0;

    virtual QByteArray ComposeByteArray() const;
    virtual bool ParseByteArray(const QByteArray& data);

    inline QSharedPointer<RelationshipMgr> GetRelationship() const { return relationship_; }
    inline void SetXmlPath(const QString& path) { xml_path_ = path; }
    inline const QString& GetXmlPath() const { return xml_path_; }

protected:
    explicit AbstractOOXmlFile(OperationMode mode = OperationMode::kCreateNew);

protected:
    QSharedPointer<RelationshipMgr> relationship_ {};
    OperationMode operation_mode_ {};
    // such as "xl/worksheets/sheet1.xml"
    QString xml_path_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_ABSTRACTOOXMLFILE_H
