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

#include "document.h"

#include <QFileInfo>
#include <QTemporaryFile>

#include "docpropsapp.h"
#include "docpropscore.h"
#include "relationshipmgr.h"
#include "sharedstring.h"
#include "style.h"
#include "utility.h"
#include "zipreader.h"
#include "zipwriter.h"

QT_BEGIN_NAMESPACE_YXLSX

void Document::Init()
{
    if (!content_type_)
        content_type_ = QSharedPointer<ContentType>::create(OperationMode::kCreateNew);

    if (!workbook_)
        workbook_ = QSharedPointer<Workbook>::create(OperationMode::kCreateNew);
}

// Explanation of the structure of an unzipped .xlsx file:
// 1. An `.xlsx` file is essentially a ZIP archive.
// 2. When unzipped, it contains multiple files and folders describing the content and structure of the Excel file, such as:
//    - `[Content_Types].xml`: Defines the content types of the file.
//    - `_rels/`: Stores relationships (e.g., between worksheets and shared strings).
//    - `docProps/`: Contains document metadata, such as title and author.
//    - `xl/`: Holds the core content of the workbook, including:
//        * `xl/workbook.xml`: Describes the workbook and lists all worksheets.
//        * `xl/worksheets/`: Stores the content of each worksheet, e.g., `sheet1.xml`.
//        * `xl/sharedStrings.xml`: Contains all shared strings to reduce redundancy.
//        * `xl/styles.xml`: Defines cell styles, such as fonts, colors, and borders.
// 3. By analyzing and parsing these files, the content of the `.xlsx` file can be extracted and manipulated efficiently.
bool Document::ParseXlsx(QIODevice* device)
{
    ZipReader zip_reader(device);
    const QStringList& file_paths { zip_reader.GetFilePath() };

    // Load the Content_Types file
    if (!file_paths.contains(QStringLiteral("[Content_Types].xml")))
        return false;

    content_type_ = QSharedPointer<ContentType>::create(OperationMode::kLoadExisting);
    content_type_->ParseByteArray(zip_reader.GetFileData(QStringLiteral("[Content_Types].xml")));

    // Load root rels file
    if (!file_paths.contains(QStringLiteral("_rels/.rels")))
        return false;
    RelationshipMgr root_rels {};
    root_rels.ReadByteArray(zip_reader.GetFileData(QStringLiteral("_rels/.rels")));

    // load core property
    QList<Relationship> core_rels { root_rels.GetPackageRelationship(QStringLiteral("/metadata/core-properties")) };
    if (!core_rels.isEmpty()) {
        // Get the core property file name if it exists.
        // In normal case, this should be "docProps/core.xml"
        const QString doc_props_core_name { core_rels[0].target };

        DocPropsCore props(OperationMode::kLoadExisting);
        props.ParseByteArray(zip_reader.GetFileData(doc_props_core_name));
        const auto prop_names { props.GetProperty() };
        for (const QString& name : prop_names)
            SetProperty(name, props.GetProperty(name));
    }

    // load app property
    QList<Relationship> rels_app { root_rels.GetDocumentRelationship(QStringLiteral("/extended-properties")) };
    if (!rels_app.isEmpty()) {
        // Get the app property file name if it exists.
        // In normal case, this should be "docProps/app.xml"
        const QString doc_props_app_Name { rels_app[0].target };

        DocPropsApp props(OperationMode::kLoadExisting);
        props.ParseByteArray(zip_reader.GetFileData(doc_props_app_Name));
        const auto prop_names { props.GetProperty() };
        for (const QString& name : prop_names)
            SetProperty(name, props.GetProperty(name));
    }

    // load workbook now, Get the workbook file path from the root rels file
    // In normal case, this should be "xl/workbook.xml"
    workbook_ = QSharedPointer<Workbook>::create(OperationMode::kLoadExisting);
    QList<Relationship> rels_xl { root_rels.GetDocumentRelationship(QStringLiteral("/officeDocument")) };
    if (rels_xl.isEmpty())
        return false;

    const QString workbook_path { rels_xl[0].target };
    const auto parts { Utility::SplitPath(workbook_path) };
    const QString& workbook_dir { parts.first() };
    const QString rel_file_path { Utility::GetRelFilePath(workbook_path) };

    workbook_->GetRelationship()->ReadByteArray(zip_reader.GetFileData(rel_file_path));
    workbook_->SetXmlPath(workbook_path);
    workbook_->ParseByteArray(zip_reader.GetFileData(workbook_path));

    // load styles
    QList<Relationship> rels_styles { workbook_->GetRelationship()->GetDocumentRelationship(QStringLiteral("/styles")) };
    if (!rels_styles.isEmpty()) {
        // In normal case this should be styles.xml which in xl
        const QString name = rels_styles[0].target;

        // dev34
        const QString path { (workbook_dir == QStringLiteral(".")) ? name : workbook_dir + QStringLiteral("/") + name };

        QSharedPointer<Style> styles { QSharedPointer<Style>::create(OperationMode::kLoadExisting) };
        styles->ParseByteArray(zip_reader.GetFileData(path));
        workbook_->GetStyle() = styles;
    }

    // load sharedStrings
    QList<Relationship> rels_sharedStrings { workbook_->GetRelationship()->GetDocumentRelationship(QStringLiteral("/sharedStrings")) };
    if (!rels_sharedStrings.isEmpty()) {
        // In normal case this should be sharedStrings.xml which in xl
        const QString name { rels_sharedStrings[0].target };
        const QString path { workbook_dir + QStringLiteral("/") + name };
        workbook_->GetSharedString()->ParseByteArray(zip_reader.GetFileData(path));
    }

    // load sheets
    int sheet_count { workbook_->GetSheetCount() };
    for (int i = 0; i != sheet_count; ++i) {
        auto sheet { workbook_->GetSheet(i) };
        const QString xml_path = sheet->GetXmlPath();
        const QString rel_path = Utility::GetRelFilePath(xml_path);
        // If the .rel file exists, load it.
        if (zip_reader.GetFilePath().contains(rel_path))
            sheet->GetRelationship()->ReadByteArray(zip_reader.GetFileData(rel_path));
        sheet->ParseByteArray(zip_reader.GetFileData(sheet->GetXmlPath()));
    }

    is_load_xlsx_ = true;
    return true;
}

bool Document::ComposeXlsx(QIODevice* device) const
{
    ZipWriter zip_writer(device);
    if (zip_writer.IsError())
        return false;

    content_type_->ClearOverride();

    DocPropsApp doc_props_app(OperationMode::kCreateNew);
    DocPropsCore doc_props_core(OperationMode::kCreateNew);

    // save worksheet xml files
    QList<QSharedPointer<AbstractSheet>> worksheets { workbook_->GetSheetByType(SheetType::kWorkSheet) };
    if (!worksheets.isEmpty())
        doc_props_app.AddHeading(QStringLiteral("Worksheets"), worksheets.size());

    for (int i = 0; i != worksheets.size(); ++i) {
        const auto& sheet = worksheets[i];
        content_type_->AddWorksheetName(QStringLiteral("sheet%1").arg(i + 1));
        doc_props_app.AddTitle(sheet->GetSheetName());

        zip_writer.AddFile(QStringLiteral("xl/worksheets/sheet%1.xml").arg(i + 1), sheet->ComposeByteArray());

        auto rel = sheet->GetRelationship();
        if (!rel->IsEmpty())
            zip_writer.AddFile(QStringLiteral("xl/worksheets/_rels/sheet%1.xml.rels").arg(i + 1), rel->WriteByteArray());
    }

    // save workbook xml file
    content_type_->AddWorkbook();
    zip_writer.AddFile(QStringLiteral("xl/workbook.xml"), workbook_->ComposeByteArray());
    zip_writer.AddFile(QStringLiteral("xl/_rels/workbook.xml.rels"), workbook_->GetRelationship()->WriteByteArray());

    // save docProps app/core xml file
    const auto doc_prop_names = document_property_hash_.keys();
    for (const QString& name : doc_prop_names) {
        doc_props_app.SetProperty(name, GetProperty(name));
        doc_props_core.SetProperty(name, GetProperty(name));
    }
    content_type_->AddDocPropApp();
    content_type_->AddDocPropCore();
    zip_writer.AddFile(QStringLiteral("docProps/app.xml"), doc_props_app.ComposeByteArray());
    zip_writer.AddFile(QStringLiteral("docProps/core.xml"), doc_props_core.ComposeByteArray());

    // save sharedStrings xml file
    if (!workbook_->GetSharedString()->IsEmpty()) {
        content_type_->AddSharedString();
        zip_writer.AddFile(QStringLiteral("xl/sharedStrings.xml"), workbook_->GetSharedString()->ComposeByteArray());
    }

    // save styles xml file
    content_type_->AddStyles();
    zip_writer.AddFile(QStringLiteral("xl/styles.xml"), workbook_->GetStyle()->ComposeByteArray());

    // save root .rels xml file
    RelationshipMgr rootrels;
    rootrels.SetDocumentRelationship(QStringLiteral("/officeDocument"), QStringLiteral("xl/workbook.xml"));
    rootrels.SetPackageRelationship(QStringLiteral("/metadata/core-properties"), QStringLiteral("docProps/core.xml"));
    rootrels.SetDocumentRelationship(QStringLiteral("/extended-properties"), QStringLiteral("docProps/app.xml"));
    zip_writer.AddFile(QStringLiteral("_rels/.rels"), rootrels.WriteByteArray());

    // save content types xml file
    zip_writer.AddFile(QStringLiteral("[Content_Types].xml"), content_type_->ComposeByteArray());

    zip_writer.Close();
    return true;
}

/*!
 * Creates a new empty xlsx document.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(QObject* parent)
    : QObject { parent }
{
    Init();
}

/*!
 * \overload
 * Try to open an existing xlsx document named \a xlsx_name.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(const QString& xlsx_name, QObject* parent)
    : QObject { parent }
    , xlsx_name_ { xlsx_name }
{
    if (xlsx_name.isEmpty()) {
        qWarning() << "Empty file name provided for the document.";
        return;
    }

    QFileInfo file_info(xlsx_name);
    if (file_info.exists() && file_info.isFile()) {
        QFile xlsx(xlsx_name);
        if (!xlsx.open(QFile::ReadOnly)) {
            qWarning() << "Failed to open the file:" << xlsx_name;
            return;
        }

        if (!ParseXlsx(&xlsx)) {
            qWarning() << "Failed to load the package for document:" << xlsx_name;
            return;
        }
    } else {
        qWarning() << "File does not exist, initializing a new document:" << xlsx_name;
    }

    Init();
}

/*!
 * Returns the value of the document's \a key property.
 * If the key is not found, returns an empty string.
 */
QString Document::GetProperty(const QString& key) const
{
    auto it = document_property_hash_.constFind(key);
    return (it != document_property_hash_.constEnd()) ? it.value() : QString();
}

/*!
        Set the document properties such as Title, Author etc.

        The method can be used to set the document properties of the Excel
        file created by Qt Xlsx. These properties are visible when you use the
        Office Button -> Prepare -> Properties option in Excel and are also
        available to external applications that read or index windows files.

        The \a property \a key that can be set are:

        \list
        \li title
        \li subject
        \li creator
        \li manager
        \li company
        \li category
        \li keywords
        \li description
        \li status
        \endlist
*/
void Document::SetProperty(const QString& key, const QString& property) { document_property_hash_[key] = property; }

/*!
 * Save current document to the filesystem. If no name specified when
 * the document constructed, a default name "book1.xlsx" will be used.
 * Returns true if saved successfully.
 */
bool Document::Save() const { return Save(xlsx_name_.isEmpty() ? kDefaultXlsxName : xlsx_name_); }

/*!
 * Saves the document to the file with the given \a name.
 * Returns true if saved successfully.
 */
bool Document::Save(const QString& xlsx_name) const
{
    QFile file(xlsx_name);
    if (!file.open(QIODevice::WriteOnly)) {
        qWarning() << "Failed to open file for writing:" << xlsx_name;
        return false;
    }

    return ComposeXlsx(&file);
}

QT_END_NAMESPACE_YXLSX
