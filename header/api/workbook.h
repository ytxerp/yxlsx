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

#ifndef YXLSX_WORKBOOK_H
#define YXLSX_WORKBOOK_H

#include <QIODevice>
#include <QSharedPointer>

#include "abstractooxmlfile.h"
#include "abstractsheet.h"
#include "definedname.h"
#include "namespace.h"
#include "sharedstring.h"
#include "style.h"
#include "worksheet.h"

QT_BEGIN_NAMESPACE_YXLSX

class Workbook final : public AbstractOOXmlFile {
    Q_DISABLE_COPY_MOVE(Workbook)

public:
    explicit Workbook(OperationMode mode);
    ~Workbook();

    QSharedPointer<AbstractSheet> AppendSheet(const QString& name = QString(), SheetType type = SheetType::kWorkSheet);
    QSharedPointer<AbstractSheet> InsertSheet(int index, const QString& name = QString(), SheetType type = SheetType::kWorkSheet);

    bool RenameSheet(int index, const QString& new_name);
    bool RenameSheet(const QString& old_name, const QString& new_name);
    bool DeleteSheet(int index);
    bool DeleteSheet(const QString& name);

    QSharedPointer<AbstractSheet> GetSheet(int index) const;
    QSharedPointer<AbstractSheet> GetSheet(const QString& sheet_name) const;

    bool SetCurrentSheet(int index);

    QSharedPointer<AbstractSheet> GetCurrentSheet() const;
    QSharedPointer<Worksheet> GetCurrentWorksheet() const;

    QStringList GetSheetName() const { return sheet_name_list_; }
    int GetSheetCount() const { return sheet_list_.count(); }

    QSharedPointer<SharedString> GetSharedString() const { return shared_string_; }
    QSharedPointer<Style> GetStyle() { return style_; }
    QList<QSharedPointer<AbstractSheet>> GetSheetByType(SheetType type) const;

private:
    void ComposeXml(QIODevice* device) const override;

    void ComposeNamespace(QXmlStreamWriter& writer) const;
    void ComposeFileVersion(QXmlStreamWriter& writer) const;
    void ComposeWorkbookProperty(QXmlStreamWriter& writer) const;
    void ComposeBookView(QXmlStreamWriter& writer) const;
    void ComposeSheet(QXmlStreamWriter& writer) const;
    void ComposeDefinedName(QXmlStreamWriter& writer) const;
    void ComposeCalcProperty(QXmlStreamWriter& writer) const;

    void SetEssentialRelationship() const;
    int GetSheetIndex(int sheet_id) const;

    bool ParseXml(QIODevice* device) override;
    void ParseSheet(QXmlStreamReader& reader);
    void ParseDefinedName(QXmlStreamReader& reader);
    void ParseBookViews(QXmlStreamReader& reader);

    QString GetXmlAttribute(const QXmlStreamAttributes& attrs, QAnyStringView key) { return attrs.hasAttribute(key) ? attrs.value(key).toString() : QString(); }
    QString ResolveFullPath(const QString& target, const QString& base_path) const;

    QSharedPointer<AbstractSheet> LoadSheet(const QString& name, int sheet_id, SheetType type = SheetType::kWorkSheet);

private:
    QSharedPointer<SharedString> shared_string_ {};
    QSharedPointer<Style> style_ {};

    QStringList sheet_name_list_ {};
    QList<DefinedName> defined_name_list_ {};
    QList<QSharedPointer<AbstractSheet>> sheet_list_ {};

    int x_window_ {};
    int y_window_ {};
    int window_width_ {};
    int window_height_ {};

    int current_sheet_index_ {};

    // Used to generate new sheet name and id
    int last_sheet_index_ {};
    int last_sheet_id_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_WORKBOOK_H
