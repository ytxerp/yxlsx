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

#include "worksheet.h"

#include <QDateTime>

#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

/*!
 * Calculates the "spans" attribute for the <row> tag in XLSX. This is an optimization for faster comparison
 * of Excel files and is not strictly required for the file's validity. The span defines the range of columns
 * occupied in a row, and the rows are grouped into blocks of 16 for this calculation.

 * The span for each block of 16 rows is stored in the row_spans_ map, with the key being the block number.
 */
void Worksheet::CalculateSpans() const
{
    row_spans_hash_.clear();
    int span_min = kExcelColumnMax + 1;
    int span_max = -1;

    for (int row_num = dimension_.TopRow(); row_num <= dimension_.BottomRow(); row_num++) {
        for (int col_num = dimension_.LeftColumn(); col_num <= dimension_.RightColumn(); col_num++) {
            if (Contains(row_num, col_num)) {
                // Update the span for the current row
                if (span_max == -1) {
                    span_min = col_num;
                    span_max = col_num;
                } else {
                    span_min = qMin(span_min, col_num);
                    span_max = qMax(span_max, col_num);
                }
            }
        }

        // After every 16 rows or at the last row, record the span and reset for the next block
        if (row_num % 16 == 0 || row_num == dimension_.BottomRow()) {
            if (span_max != -1) {
                row_spans_hash_[row_num / 16] = QStringLiteral("%1:%2").arg(span_min).arg(span_max);
                span_min = kExcelColumnMax + 1;
                span_max = -1;
            }
        }
    }
}

QString Worksheet::ComposeDimension() const
{
    if (!dimension_.CheckValid())
        return QStringLiteral("A1");
    else
        return dimension_.ComposeDimension();
}

/*!
 * Updates the worksheet's dimension based on the given row and column.
 * Returns true if the dimension was successfully updated, false otherwise.
 */
bool Worksheet::UpdateDimension(int row, int col)
{
    if (!Utility::CheckCoordinateValid(row, col))
        return false;

    Q_ASSERT_X(row > 0, "UpdateDimension", "Row index must be 1 or greater.");
    Q_ASSERT_X(col > 0, "UpdateDimension", "Column index must be 1 or greater.");

    // Update the dimension if the new row or column extends beyond current bounds.
    dimension_.SetTopRow(row);
    dimension_.SetBottomRow(row);
    dimension_.SetLeftColumn(col);
    dimension_.SetRightColumn(col);

    return true;
}

/*!
 * \internal
 */
Worksheet::Worksheet(const QString& sheet_name, int sheet_id, QSharedPointer<SharedString> shared_strings, SheetType sheet_type)
    : AbstractSheet { sheet_name, sheet_id, sheet_type }
    , shared_string_ { shared_strings }
{
}

/*!
 * Destroys this workssheet.
 */
Worksheet::~Worksheet() { }

/*!
 * Writes \a data to the cell at the given \a row and \a column.
 * Both \a row and \a column are 1-indexed.
 *
 * Returns true if the write operation is successful.
 */
bool Worksheet::Write(int row, int column, const QVariant& data)
{
    if (!Utility::CheckCoordinateValid(row, column))
        return false;

    if (data.isNull() || !UpdateDimension(row, column)) {
        return false;
    }

    CellType cell_type { DetermineCellType(data) };
    if (cell_type == CellType::kUnknown)
        return false;

    auto cell { QSharedPointer<Cell>::create(data, cell_type) };

    if (cell_type == CellType::kSharedString)
        shared_string_->SetSharedString(data.toString(), row, column);

    WriteMatrix(row, column, cell);
    return true;
}

CellType Worksheet::DetermineCellType(const QVariant& value) const
{
    if (!value.isValid() || value.isNull()) {
        qWarning() << "Invalid or null QVariant provided.";
        return CellType::kUnknown;
    }

    const int kTypeID { value.typeId() };
    for (const auto& [meta_type, cell_type] : type_map) {
        if (kTypeID == meta_type) {
            return cell_type;
        }
    }

    qWarning() << "Unsupported QVariant type:" << value.typeName();
    return CellType::kUnknown;
}

/*!
 * \overload
 * Writes \a data to the cell at the given \a coordinate.
 * Both row and column in the \a coordinate are 1-indexed.
 * Returns true if the operation is successful.
 */
bool Worksheet::Write(const Coordinate& coordinate, const QVariant& data)
{
    // Ensure the coordinate is valid before proceeding
    if (!Coordinate::CheckValid(coordinate))
        return false;

    // Delegate to the row-column-based Write method
    return Write(coordinate.Row(), coordinate.Column(), data);
}

/*!
 * \overload
 * Returns the value stored in the cell at the specified \a coordinate.
 * If the \a coordinate is invalid, an empty QVariant is returned.
 */
QVariant Worksheet::Read(const Coordinate& coordinate) const
{
    // Check if the provided coordinate is valid
    if (!Coordinate::CheckValid(coordinate))
        return QVariant();

    // Delegate to the row and column-based Read function
    return Read(coordinate.Row(), coordinate.Column());
}

/*!
 * Returns the value stored in the cell located at the specified \a row and \a column.
 * If the cell does not exist, an empty QVariant is returned.
 */
QVariant Worksheet::Read(int row, int column) const
{
    // Retrieve the cell at the given position
    auto cell { ReadMatrix(row, column) };

    // Return the cell's value if it exists; otherwise, return an empty QVariant
    return cell ? cell->value : QVariant();
}

/*!
        Write a empty cell (\a row, \a column) with the \a format.
        Returns true on success.
 */
bool Worksheet::WriteBlank(int row, int column)
{
    // Note: NumberType with an invalid QVariant value means blank.
    auto cell { QSharedPointer<Cell>::create(QVariant {}, CellType::kNumber) };
    WriteMatrix(row, column, cell);
    return true;
}

/*!
 * \internal
 */
void Worksheet::ComposeXml(QIODevice* device) const
{
    relationship_->Clear();

    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QLatin1String("1.0"), true);
    writer.writeStartElement(QLatin1String("worksheet"));
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QLatin1String("xmlns:r"), QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    writer.writeStartElement(QLatin1String("dimension"));
    writer.writeAttribute(QLatin1String("ref"), ComposeDimension());
    writer.writeEndElement(); // dimension

    writer.writeStartElement(QLatin1String("sheetViews"));
    writer.writeStartElement(QLatin1String("sheetView"));
    writer.writeAttribute(QLatin1String("workbookViewId"), QLatin1String("0"));
    writer.writeEndElement(); // sheetView
    writer.writeEndElement(); // sheetViews

    writer.writeStartElement(QLatin1String("sheetFormatPr"));
    writer.writeAttribute(QLatin1String("defaultRowHeight"), QString::number(sheet_format_props_.default_row_height));
    writer.writeAttribute(QLatin1String("defaultColWidth"), QString::number(sheet_format_props_.default_col_width));
    writer.writeEndElement(); // sheetFormatPr

    writer.writeStartElement(QLatin1String("sheetData"));
    if (dimension_.CheckValid())
        ComposeSheet(writer);
    writer.writeEndElement(); // sheetData

    writer.writeEndElement(); // worksheet
    writer.writeEndDocument();
}

void Worksheet::ComposeSheet(QXmlStreamWriter& writer) const
{
    // Ensure spans are calculated before writing
    CalculateSpans();

    const int bottom { dimension_.BottomRow() };
    const int left { dimension_.LeftColumn() };

    for (int row = dimension_.TopRow(); row <= bottom; ++row) {
        // Determine the span attribute for the current row
        QString span {};
        int span_index = (row - 1) / 16;
        if (auto span_it = row_spans_hash_.constFind(span_index); span_it != row_spans_hash_.constEnd()) {
            span = span_it.value();
        }

        // Start the row element
        writer.writeStartElement(QLatin1String("row"));
        writer.writeAttribute(QLatin1String("r"), QString::number(row));
        if (!span.isEmpty()) {
            writer.writeAttribute(QLatin1String("spans"), span);
        }

        // Write cell data for the row if it contains any cells
        for (auto it = matrix_.lowerBound({ row, left }); it != matrix_.constEnd(); ++it) {
            const auto& coord { it.key() };

            // If the current row number is greater than the target row number, no need to continue
            if (coord.first > row) {
                break;
            }

            // Process only the cells in the target row
            // Ensure the cell pointer is valid and contains valid data before writing
            if (it.value() && it.value()->value.isValid()) {
                ComposeCell(writer, row, coord.second, *it);
            }
        }

        // End the row element
        writer.writeEndElement();
    }
}

void Worksheet::ComposeCell(QXmlStreamWriter& writer, int row, int col, const QSharedPointer<Cell>& cell) const
{
    Q_ASSERT(cell);
    if (!cell->value.isValid() || cell->value.isNull())
        return;

    // This is the innermost loop so efficiency is important.
    QString coord = Utility::ComposeCoordinate(row, col);

    writer.writeStartElement(QLatin1String("c"));
    writer.writeAttribute(QLatin1String("r"), coord);

    switch (cell->type) {
    case CellType::kSharedString: { // 's'
        int shared_string_index { shared_string_->GetSharedStringIndex(cell->value.toString()) };
        writer.writeAttribute(QLatin1String("t"), QLatin1String("s"));
        writer.writeTextElement(QLatin1String("v"), QString::number(shared_string_index));
        break;
    }
    case CellType::kNumber: { // 'n'
        writer.writeAttribute(QLatin1String("t"), QLatin1String("n"));
        writer.writeTextElement(QLatin1String("v"), QString::number(cell->value.toDouble(), 'g', 15));
        break;
    }
    case CellType::kBoolean: { // 'b'
        writer.writeAttribute(QLatin1String("t"), QLatin1String("b"));
        writer.writeTextElement(QLatin1String("v"), cell->value.toBool() ? QLatin1String("1") : QLatin1String("0"));
        break;
    }
    case CellType::kDate: {
        writer.writeAttribute(QLatin1String("t"), QLatin1String("d"));
        writer.writeTextElement(QLatin1String("v"), cell->value.toDateTime().toString(Qt::ISODate));
        break;
    }
    default:
        qDebug() << "Unknown CellType encountered!";
        break;
    }

    writer.writeEndElement();
}

void Worksheet::ParseSheet(QXmlStreamReader& reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("sheetData"));

    int row = 0;
    int column = 0;

    while (!reader.atEnd() && !(reader.name() == QLatin1String("sheetData") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        if (!reader.readNextStartElement())
            continue;

        const auto name { reader.name() };

        if (name == QStringLiteral("row")) { // Row
            QXmlStreamAttributes attributes = reader.attributes();

            if (attributes.hasAttribute(QLatin1String("r")))
                row = attributes.value(QLatin1String("r")).toInt();
            else
                ++row;

            column = 0;

        } else if (name == QStringLiteral("c")) // Cell
        {
            ProcessCell(reader, row, column);
        }

        if (reader.hasError()) {
            qWarning() << "XML Parsing Error in sheetData:" << reader.errorString();
        }
    }
}

void Worksheet::ProcessCell(QXmlStreamReader& reader, int row, int& column)
{
    Q_ASSERT(reader.name() == QStringLiteral("c"));

    QXmlStreamAttributes attributes { reader.attributes() };
    QString cell_reference = attributes.value(QLatin1String("r")).toString();

    // Determine cell position
    Coordinate coord(cell_reference);
    if (cell_reference.isEmpty()) {
        coord.SetRow(row);
        coord.SetColumn(++column);
    } else {
        column = coord.Column();
    }

    CellType cell_type {};

    if (attributes.hasAttribute(QLatin1String("t"))) // Type
    {
        const auto type = attributes.value(QLatin1String("t"));
        if (type == QLatin1String("s")) // Shared string
        {
            cell_type = CellType::kSharedString;
        } else if (type == QLatin1String("b")) // Boolean
        {
            cell_type = CellType::kBoolean;
        } else if (type == QLatin1String("d")) // Date
        {
            cell_type = CellType::kDate;
        } else if (type == QLatin1String("n")) // Number
        {
            cell_type = CellType::kNumber;
        } else {
            cell_type = CellType::kNumber;
        }
    }

    auto cell { QSharedPointer<Cell>::create(QVariant {}, cell_type) };

    // Parse the cell's sub-elements
    while (!reader.atEnd() && !(reader.name() == QStringLiteral("c") && reader.tokenType() == QXmlStreamReader::EndElement)) {
        if (reader.readNextStartElement() && reader.name() == QStringLiteral("v")) {
            // Parse cell value
            QString value = reader.readElementText();
            cell->value = ParseCellValue(value, cell_type, row, column);
        }
    }

    if (reader.hasError()) {
        qWarning() << "XML Parsing Error in <c> element:" << reader.errorString();
    }

    // Write the cell to the matrix
    WriteMatrix(coord.Row(), coord.Column(), cell);
}

QVariant Worksheet::ParseCellValue(const QString& value, CellType cell_type, int row, int column)
{
    switch (cell_type) {
    case CellType::kSharedString: {
        int shared_string_index = value.toInt();
        shared_string_->IncrementReference(shared_string_index, row, column);
        return shared_string_->GetSharedString(shared_string_index);
    }
    case CellType::kBoolean:
        return QVariant(value == QStringLiteral("1") || value.toLower() == QStringLiteral("true"));
    case CellType::kDate:
        return QVariant::fromValue(QDateTime::fromString(value, Qt::ISODate));
    case CellType::kNumber:
        return QVariant(value.toDouble());
    default:
        qWarning() << "Unsupported cell type, returning original value:" << value;
        return value;
    }
}

bool Worksheet::ParseXml(QIODevice* device)
{
    if (!device || !device->isOpen()) {
        qWarning() << "Invalid or unopened QIODevice.";
        return false;
    }

    QXmlStreamReader reader(device);
    QStringView name {};

    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        name = reader.name();

        if (name == QStringLiteral("dimension")) {
            QXmlStreamAttributes attributes = reader.attributes();
            QString range = attributes.value(QLatin1String("ref")).toString();
            dimension_ = Dimension(range);
        } else if (name == QStringLiteral("sheetData")) {
            ParseSheet(reader);
        }
    }

    if (reader.hasError()) {
        qWarning() << "XML Parsing Error:" << reader.errorString();
        return false;
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
