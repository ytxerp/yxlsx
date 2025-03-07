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

#ifndef YXLSX_DOCPROPSAPP_H
#define YXLSX_DOCPROPSAPP_H

#include <QHash>
#include <QIODevice>
#include <QList>

#include "abstractooxmlfile.h"
#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

class DocPropsApp final : public AbstractOOXmlFile {
public:
    explicit DocPropsApp(OperationMode mode);

    void ComposeXml(QIODevice* device) const override;
    bool ParseXml(QIODevice* device) override;

    bool SetProperty(const QString& name, const QString& value);

    inline void AddTitle(const QString& title) { title_list_.append(title); }
    inline void AddHeading(const QString& name, int value) { heading_list_.append({ name, value }); }

    inline QString GetProperty(const QString& name) const { return property_hash_.value(name, QString()); }
    inline QStringList GetProperty() const { return property_hash_.keys(); }

private:
    QStringList title_list_ {};
    QList<std::pair<QString, int>> heading_list_ {};
    QHash<QString, QString> property_hash_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_DOCPROPSAPP_H
