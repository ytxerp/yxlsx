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

#ifndef YXLSX_DOCUMENT_H
#define YXLSX_DOCUMENT_H

#include "contenttype.h"
#include "workbook.h"

QT_BEGIN_NAMESPACE_YXLSX

class Document final : public QObject {
public:
    explicit Document(QObject* parent = nullptr);
    explicit Document(const QString& xlsx_name, QObject* parent = nullptr);
    ~Document() = default;

    QString GetProperty(const QString& key) const;
    void SetProperty(const QString& key, const QString& property);

    bool Save() const;
    bool Save(const QString& xlsx_name) const;

    bool IsLoadXlsx() const { return is_load_xlsx_; }
    QSharedPointer<Workbook> GetWorkbook() const { return workbook_; }
    QStringList GetProperty() const { return document_property_hash_.keys(); }

private:
    void Init();
    bool ParseXlsx(QIODevice* device);
    bool ComposeXlsx(QIODevice* device) const;

private:
    bool is_load_xlsx_ { false };

    QString xlsx_name_ {}; // name of the .xlsx file

    QHash<QString, QString> document_property_hash_ {}; // core, app and custom properties
    QSharedPointer<Workbook> workbook_ {};
    QSharedPointer<ContentType> content_type_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_DOCUMENT_H
