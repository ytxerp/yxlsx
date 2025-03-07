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

#include "workbook.h"

#include <QDir>

#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

Workbook::Workbook(OperationMode mode)
    : shared_string_ { QSharedPointer<SharedString>::create(mode) }
    , style_ { QSharedPointer<Style>::create(mode) }
    , x_window_ { 240 }
    , y_window_ { 15 }
    , window_width_ { 16095 }
    , window_height_ { 9660 }
{
}
Workbook::~Workbook() { }

QSharedPointer<AbstractSheet> Workbook::AppendSheet(const QString& name, SheetType type) { return InsertSheet(sheet_list_.size(), name, type); }

/*!
 * \internal
 * Used only when load the xlsx file!!
 */
QSharedPointer<AbstractSheet> Workbook::LoadSheet(const QString& name, int sheet_id, SheetType type)
{
    // Ensure the last_sheet_id_ is updated
    last_sheet_id_ = std::max(last_sheet_id_, sheet_id);

    // Check if the sheet type is supported
    if (type != SheetType::kWorkSheet) {
        qWarning("Unsupported sheet type: %d", static_cast<int>(type));
        return {}; // Return immediately for unsupported types
    }

    // Create the appropriate sheet instance
    auto sheet { QSharedPointer<Worksheet>::create(name, sheet_id, shared_string_, type) };

    // Store the sheet and its name in the workbook's containers
    sheet_list_.emplaceBack(sheet);
    sheet_name_list_.emplaceBack(name);

    return sheet; // Return the created sheet
}

QSharedPointer<AbstractSheet> Workbook::InsertSheet(int index, const QString& name, SheetType type)
{
    // Validate index
    if (index < 0 || index > sheet_list_.size()) {
        qWarning("Invalid index for sheet insertion: %d", index);
        return {};
    }

    // Generate a safe and unique sheet name
    QString sheet_name { Utility::GenerateSheetName(sheet_name_list_, name, last_sheet_index_) };

    // Update sheet ID and create the sheet
    ++last_sheet_id_;
    auto sheet { QSharedPointer<Worksheet>::create(sheet_name, last_sheet_id_, shared_string_, type) };

    // Insert the sheet and name into containers
    sheet_list_.insert(index, sheet);
    sheet_name_list_.insert(index, sheet_name);

    // Update the active sheet index
    current_sheet_index_ = index;

    return sheet;
}

bool Workbook::SetCurrentSheet(int index)
{
    // Validate the index range
    if (index < 0 || index >= sheet_list_.size()) {
        qWarning() << "Invalid sheet index:" << index << ". No changes made.";
        return false;
    }

    // Set the active sheet index
    current_sheet_index_ = index;
    return true;
}

QSharedPointer<AbstractSheet> Workbook::GetSheet(const QString& sheet_name) const { return GetSheet(GetSheetName().indexOf(sheet_name)); }

QSharedPointer<AbstractSheet> Workbook::GetCurrentSheet() const
{
    if (sheet_list_.isEmpty())
        const_cast<Workbook*>(this)->AppendSheet();

    return sheet_list_.at(current_sheet_index_);
}

QSharedPointer<Worksheet> Workbook::GetCurrentWorksheet() const
{
    auto current_sheet { GetCurrentSheet() };
    if (current_sheet && current_sheet->GetSheetType() == SheetType::kWorkSheet) {
        return current_sheet.dynamicCast<Worksheet>();
    }

    return {};
}

/*!
 * Rename the worksheet at the \a index to \a new_name.
 */
bool Workbook::RenameSheet(int index, const QString& new_name)
{
    // Validate index
    if (index < 0 || index >= sheet_list_.size()) {
        qWarning() << "Invalid sheet index:" << index;
        return false;
    }

    QString safe_name { Utility::GenerateSheetName(sheet_name_list_, new_name, last_sheet_index_) };

    sheet_list_[index]->SetSheetName(safe_name);
    sheet_name_list_[index] = safe_name;
    return true;
}

bool Workbook::RenameSheet(const QString& old_name, const QString& new_name)
{
    auto index { sheet_name_list_.indexOf(old_name) };
    return RenameSheet(index, new_name);
}

/*!
 * Remove the worksheet at pos \a index.
 */
bool Workbook::DeleteSheet(int index)
{
    // Ensure there is more than one sheet to delete
    if (sheet_list_.size() <= 1) {
        qWarning() << "Cannot delete the only remaining sheet.";
        return false;
    }

    // Validate index
    if (index < 0 || index >= sheet_list_.size()) {
        qWarning() << "Invalid sheet index:" << index;
        return false;
    }

    sheet_list_.removeAt(index);
    sheet_name_list_.removeAt(index);

    // Adjust active sheet index if necessary
    if (current_sheet_index_ >= index) {
        current_sheet_index_ = qMax(0, current_sheet_index_ - 1);
    }

    return true;
}

bool Workbook::DeleteSheet(const QString& name)
{
    auto index { sheet_name_list_.indexOf(name) };
    return DeleteSheet(index);
}

/*!
 * Returns the sheet object at index \a sheetIndex.
 */
QSharedPointer<AbstractSheet> Workbook::GetSheet(int index) const
{
    if (index < 0 || index >= sheet_list_.size())
        return {};

    return sheet_list_.at(index);
}

/*!
 * \internal
 */
QList<QSharedPointer<AbstractSheet>> Workbook::GetSheetByType(SheetType type) const
{
    QList<QSharedPointer<AbstractSheet>> list {};
    list.reserve(sheet_list_.size());

    for (const auto& sheet : sheet_list_) {
        if (sheet->GetSheetType() == type) {
            list.append(sheet);
        }
    }

    return list;
}

void Workbook::ComposeXml(QIODevice* device) const
{
    // Ensure relationships are cleared and at least one sheet exists
    relationship_->Clear();
    if (sheet_list_.isEmpty())
        const_cast<Workbook*>(this)->AppendSheet();

    QXmlStreamWriter writer(device);
    writer.writeStartDocument(QLatin1String("1.0"), true);

    // Start workbook root element
    writer.writeStartElement(QLatin1String("workbook"));
    ComposeNamespace(writer);

    // Write file and workbook properties
    ComposeFileVersion(writer);
    ComposeWorkbookProperty(writer);

    // Write book views (window settings)
    ComposeBookView(writer);

    // Write sheets information
    ComposeSheet(writer);

    // Write defined names if available
    if (!defined_name_list_.isEmpty())
        ComposeDefinedName(writer);

    // Write calculation properties
    ComposeCalcProperty(writer);

    // End workbook element and document
    writer.writeEndElement(); // workbook
    writer.writeEndDocument();

    // Add essential relationships
    SetEssentialRelationship();
}

// Helper: Write XML namespaces
void Workbook::ComposeNamespace(QXmlStreamWriter& writer) const
{
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QLatin1String("xmlns:r"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
}

// Helper: Write file version
void Workbook::ComposeFileVersion(QXmlStreamWriter& writer) const
{
    writer.writeEmptyElement(QLatin1String("fileVersion"));
    writer.writeAttribute(QLatin1String("appName"), QLatin1String("xl"));
    writer.writeAttribute(QLatin1String("lastEdited"), QLatin1String("4"));
    writer.writeAttribute(QLatin1String("lowestEdited"), QLatin1String("4"));
    writer.writeAttribute(QLatin1String("rupBuild"), QLatin1String("4505"));
}

// Helper: Write workbook properties
void Workbook::ComposeWorkbookProperty(QXmlStreamWriter& writer) const
{
    writer.writeEmptyElement(QLatin1String("workbookPr"));
    writer.writeAttribute(QLatin1String("defaultThemeVersion"), QLatin1String("124226"));
}

// Helper: Write book views
void Workbook::ComposeBookView(QXmlStreamWriter& writer) const
{
    writer.writeStartElement(QLatin1String("bookViews"));
    writer.writeEmptyElement(QLatin1String("workbookView"));
    writer.writeAttribute(QLatin1String("xWindow"), QString::number(x_window_));
    writer.writeAttribute(QLatin1String("yWindow"), QString::number(y_window_));
    writer.writeAttribute(QLatin1String("windowWidth"), QString::number(window_width_));
    writer.writeAttribute(QLatin1String("windowHeight"), QString::number(window_height_));
    if (current_sheet_index_ > 0)
        writer.writeAttribute(QLatin1String("activeTab"), QString::number(current_sheet_index_));
    writer.writeEndElement(); // bookViews
}

// Helper: Write sheets
void Workbook::ComposeSheet(QXmlStreamWriter& writer) const
{
    writer.writeStartElement(QLatin1String("sheets"));

    int worksheet_index { 0 };

    for (const auto& sheet : sheet_list_) {
        writer.writeEmptyElement(QLatin1String("sheet"));
        writer.writeAttribute(QLatin1String("name"), sheet->GetSheetName());
        writer.writeAttribute(QLatin1String("sheetId"), QString::number(sheet->GetSheetId()));

        if (sheet->GetSheetType() == SheetType::kWorkSheet) {
            relationship_->SetDocumentRelationship(QStringLiteral("/worksheet"), QStringLiteral("worksheets/sheet%1.xml").arg(++worksheet_index));
        }

        writer.writeAttribute(QLatin1String("r:id"), QStringLiteral("rId%1").arg(relationship_->Count()));
    }

    writer.writeEndElement(); // sheets
}

// Helper: Write defined names
void Workbook::ComposeDefinedName(QXmlStreamWriter& writer) const
{
    writer.writeStartElement(QLatin1String("definedNames"));

    for (const DefinedName& data : defined_name_list_) {
        writer.writeStartElement(QLatin1String("definedName"));
        writer.writeAttribute(QLatin1String("name"), data.name);

        if (!data.comment.isEmpty())
            writer.writeAttribute(QLatin1String("comment"), data.comment);

        if (data.sheet_id != -1) {
            int localSheetId = GetSheetIndex(data.sheet_id);
            if (localSheetId != -1)
                writer.writeAttribute(QLatin1String("localSheetId"), QString::number(localSheetId));
        }

        writer.writeCharacters(data.formula);
        writer.writeEndElement(); // definedName
    }

    writer.writeEndElement(); // definedNames
}

// Helper: Get local sheet index
int Workbook::GetSheetIndex(int sheet_id) const
{
    const auto size { sheet_list_.size() };

    for (qsizetype i = 0; i != size; ++i) {
        if (sheet_list_.at(i)->GetSheetId() == sheet_id)
            return i;
    }

    return -1; // Not found
}

// Helper: Write calculation properties
void Workbook::ComposeCalcProperty(QXmlStreamWriter& writer) const
{
    writer.writeStartElement(QLatin1String("calcPr"));
    writer.writeAttribute(QLatin1String("calcId"), QLatin1String("124519"));
    writer.writeEndElement(); // calcPr
}

// Helper: Add essential relationships
void Workbook::SetEssentialRelationship() const
{
    relationship_->SetDocumentRelationship(QStringLiteral("/theme"), QStringLiteral("theme/theme1.xml"));
    relationship_->SetDocumentRelationship(QStringLiteral("/styles"), QStringLiteral("styles.xml"));

    if (!GetSharedString()->IsEmpty())
        relationship_->SetDocumentRelationship(QStringLiteral("/sharedStrings"), QStringLiteral("sharedStrings.xml"));
}

bool Workbook::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);
    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        const auto element_name { reader.name() };

        // Parse "sheet" elements
        if (element_name == QStringLiteral("sheet")) {
            ParseSheet(reader);
        }
        // Parse "bookviews" elements
        else if (element_name == QStringLiteral("bookviews")) {
            ParseBookViews(reader);
        }
        // Parse "definedName" elements
        else if (element_name == QStringLiteral("definedName")) {
            ParseDefinedName(reader);
        }
    }

    if (reader.hasError()) {
        qWarning() << "XML Parsing Error:" << reader.errorString();
        return false;
    }

    return true;
}

// Helper function to parse <bookviews> and <workbookView> elements
void Workbook::ParseBookViews(QXmlStreamReader& reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("bookviews"));

    while (!reader.atEnd() && !(reader.name() == QStringLiteral("bookviews") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        if (reader.readNextStartElement() && reader.name() == QStringLiteral("workbookView")) {
            const auto attributes { reader.attributes() };

            x_window_ = GetXmlAttribute(attributes, QLatin1String("xWindow")).toInt();
            y_window_ = GetXmlAttribute(attributes, QLatin1String("yWindow")).toInt();
            window_width_ = GetXmlAttribute(attributes, QLatin1String("windowWidth")).toInt();
            window_height_ = GetXmlAttribute(attributes, QLatin1String("windowHeight")).toInt();
            current_sheet_index_ = GetXmlAttribute(attributes, QLatin1String("activeTab")).toInt();
        }
    }
}

// Helper function to parse <sheet> elements
void Workbook::ParseSheet(QXmlStreamReader& reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("sheet"));

    const auto attributes { reader.attributes() };

    const QString name { GetXmlAttribute(attributes, QLatin1String("name")) };
    const int sheet_id { GetXmlAttribute(attributes, QLatin1String("sheetId")).toInt() };
    const QString r_id { GetXmlAttribute(attributes, QLatin1String("r:id")) };

    // Retrieve relationship information
    const auto relationship { relationship_ ? relationship_->GetRelationshipByID(r_id) : Relationship {} };
    if (relationship.target.isEmpty()) {
        qWarning() << "Failed to resolve relationship for sheet ID:" << sheet_id << "Name:" << name;
        return;
    }

    // Determine sheet type based on the relationship type
    SheetType type = SheetType::kWorkSheet;
    if (!relationship.type.endsWith(QStringLiteral("/worksheet"))) {
        qWarning() << "Unknown sheet type:" << relationship.type;
    }

    auto sheet { LoadSheet(name, sheet_id, type) };
    if (!sheet) {
        qWarning() << "Failed to load sheet for sheet ID:" << sheet_id << "Name:" << name;
        return;
    }

    // Resolve the full XML path
    const QString full_path { ResolveFullPath(relationship.target, GetXmlPath()) };
    sheet->SetXmlPath(full_path);
}

// Helper function to parse <definedName> elements
void Workbook::ParseDefinedName(QXmlStreamReader& reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("definedName"));

    const auto attributes { reader.attributes() };
    DefinedName data;

    data.name = GetXmlAttribute(attributes, QLatin1String("name"));
    data.comment = GetXmlAttribute(attributes, QLatin1String("comment"));

    if (attributes.hasAttribute(QLatin1String("localSheetId"))) {
        int localId = attributes.value(QLatin1String("localSheetId")).toInt();
        data.sheet_id = sheet_list_.at(localId)->GetSheetId();
    }

    data.formula = reader.readElementText();
    defined_name_list_.append(data);
}

QString Workbook::ResolveFullPath(const QString& target, const QString& base_path) const
{
    if (target.startsWith(QStringLiteral("/"))) {
        return QDir::cleanPath(target.mid(1)); // Absolute path
    } else {
        const auto parts { Utility::SplitPath(base_path) };
        return QDir::cleanPath(parts.first() + QStringLiteral("/") + target); // Relative path
    }
}

QT_END_NAMESPACE_YXLSX
