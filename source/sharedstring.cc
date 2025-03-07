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

#include "sharedstring.h"

#include <QDebug>

#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

SharedString::SharedString(OperationMode mode)
    : AbstractOOXmlFile { mode }
{
}

int SharedString::SetSharedString(const QString& string, int row, int column)
{
    // If the string is not in the list, append it
    if (!string_list_.contains(string)) {
        int index = string_list_.size();
        string_list_.append(string);
        string_index_hash_.insert(string, index);
    }

    // Insert or update the cell coordinates for the shared string
    string_coordinate_hash_[string].insert({ row, column });

    return string_index_hash_.value(string);
}

void SharedString::IncrementReference(int index, int row, int column)
{
    if (index < 0 || index >= string_list_.size()) {
        qDebug("SharedStrings: invalid index");
        return;
    }

    string_coordinate_hash_[string_list_.at(index)].insert({ row, column });
}

QSet<std::pair<int, int>> SharedString::RemoveSharedString(const QString& string, int row, int column)
{
    const int index { string_index_hash_.value(string, -1) };
    if (index == -1) {
        qDebug() << "String not found:" << string;
        return {};
    }

    if (!string_coordinate_hash_[string].remove({ row, column })) {
        qDebug() << "Row not found:" << row;
        return {};
    }

    QSet<std::pair<int, int>> affected_coord {};

    if (string_coordinate_hash_.value(string).isEmpty()) {
        for (int i = index + 1; i != string_list_.size(); ++i) {
            const QString& target_string { string_list_.at(i) };
            const auto& target_set { string_coordinate_hash_.value(target_string) };

            string_index_hash_[target_string] -= 1;

            for (const auto& pair : target_set)
                affected_coord += pair;
        }

        string_coordinate_hash_.remove(string);
        string_index_hash_.remove(string);
        string_list_.removeAt(index);
    }

    return affected_coord;
}

QString SharedString::GetSharedString(int index) const
{
    if (index < string_list_.count() && index >= 0)
        return string_list_[index];
    return {};
}

void SharedString::ComposeXml(QIODevice* device) const
{
    QXmlStreamWriter writer(device);

    if (string_list_.size() != string_index_hash_.size()) {
        qDebug("Warning: Duplicated string items exist in shared_string_list_.");
    }

    // Initialize XML document
    writer.writeStartDocument(QLatin1String("1.0"), true);

    // Calculate the total reference count
    int total_count = std::accumulate(
        string_coordinate_hash_.cbegin(), string_coordinate_hash_.cend(), 0, [](int sum, const QSet<QPair<int, int>>& set) { return sum + set.size(); });

    // Write root element <sst>
    writer.writeStartElement(QLatin1String("sst"));
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QLatin1String("count"), QString::number(total_count));
    writer.writeAttribute(QLatin1String("uniqueCount"), QString::number(string_list_.size()));

    // Write each shared string
    for (const QString& string : string_list_) {
        writer.writeStartElement(QLatin1String("si")); // <si>
        writer.writeStartElement(QLatin1String("t")); // <t>

        // Add xml:space attribute if needed
        if (Utility::IsSpaceReserveNeeded(string)) {
            writer.writeAttribute(QLatin1String("xml:space"), QLatin1String("preserve"));
        }

        writer.writeCharacters(string); // Write string content
        writer.writeEndElement(); // </t>
        writer.writeEndElement(); // </si>
    }

    writer.writeEndElement(); // </sst>
    writer.writeEndDocument(); // End document
}

void SharedString::ParseSharedString(QXmlStreamReader& reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("si"));

    QString string {};

    while (!reader.atEnd() && !(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QStringLiteral("si"))) {
        if (reader.readNextStartElement() && reader.name() == QStringLiteral("t")) {
            string += reader.readElementText();
        }
    }

    if (!string.isEmpty()) {
        const auto index { string_list_.size() };
        string_list_.append(string);
        string_index_hash_.insert(string, index);
    }
}

bool SharedString::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);
    qsizetype unique_count = 0;
    bool has_unique_count_attr = true;

    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        if (reader.name() == QStringLiteral("sst")) {
            // Process root <sst> node, check for 'uniqueCount' attribute
            const auto attributes { reader.attributes() };
            if ((has_unique_count_attr = attributes.hasAttribute(QLatin1String("uniqueCount")))) {
                bool ok = false;
                unique_count = attributes.value(QLatin1String("uniqueCount")).toInt(&ok);
                if (!ok) {
                    qDebug("Error: Failed to parse 'uniqueCount' attribute.");
                    return false;
                }
            }
        } else if (reader.name() == QStringLiteral("si")) {
            // Process shared string node <si>
            ParseSharedString(reader);
        }
    }

    // Handle potential XML errors
    if (reader.hasError()) {
        qDebug() << "Error: Failed to read XML:" << reader.errorString();
        return false;
    }

    // Validate uniqueCount if attribute was present
    if (has_unique_count_attr && string_list_.size() != unique_count) {
        qDebug("Error: Shared string count mismatch. Expected %lld, found %lld.", unique_count, string_list_.size());
        return false;
    }

    if (string_list_.size() != string_index_hash_.size()) {
        qDebug("Warning: Duplicated items exist in shared string table.");
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
