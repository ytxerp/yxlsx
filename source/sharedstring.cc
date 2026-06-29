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

YXLSX_BEGIN_NAMESPACE

SharedString::SharedString(OperationMode mode)
    : AbstractOOXmlFile { mode }
{
}

void SharedString::SetSharedString(const QString& string)
{
    auto it = string_index_hash_.find(string);

    // If string does not exist, create new entry
    if (it == string_index_hash_.end()) {
        int index = string_list_.size();

        string_list_.append(string);
        it = string_index_hash_.insert(string, index);
    }

    // Optional usage tracking
    string_count_hash_[string] += 1;
}

void SharedString::IncrementReference(int index)
{
    if (index < 0 || index >= string_list_.size()) {
        qDebug("SharedStrings: invalid index");
        return;
    }

    string_count_hash_[string_list_.at(index)] += 1;
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
    int total_count { 0 };
    for (auto it = string_count_hash_.cbegin(); it != string_count_hash_.cend(); ++it) {
        total_count += it.value();
    }

    // Write root element <sst>
    writer.writeStartElement(QLatin1String("sst"));
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QLatin1String("count"), QString::number(total_count));
    writer.writeAttribute(QLatin1String("uniqueCount"), QString::number(string_list_.size()));

    // Write each shared string
    for (const QString& string : string_list_) {
        writer.writeStartElement(QLatin1String("si"));
        writer.writeStartElement(QLatin1String("t"));

        // Add xml:space attribute if needed
        if (Utility::IsSpaceReserveNeeded(string)) {
            writer.writeAttribute(QLatin1String("http://www.w3.org/XML/1998/namespace"), QLatin1String("space"), QLatin1String("preserve"));
        }
        writer.writeCharacters(string);
        writer.writeEndElement(); // </t>
        writer.writeEndElement(); // </si>
    }

    writer.writeEndElement(); // </sst>
    writer.writeEndDocument();
}

void SharedString::ParseSharedString(QXmlStreamReader& reader)
{
    Q_ASSERT(reader.name() == QStringLiteral("si"));

    QString string {};

    while (reader.readNextStartElement()) {
        if (reader.name() == QStringLiteral("t")) {
            string += reader.readElementText(QXmlStreamReader::IncludeChildElements);
        } else if (reader.name() == QStringLiteral("r")) {
            // rich text run → drill down automatically
            while (reader.readNextStartElement()) {
                if (reader.name() == QStringLiteral("t")) {
                    string += reader.readElementText();
                } else {
                    reader.skipCurrentElement();
                }
            }
        } else {
            reader.skipCurrentElement();
        }
    }

    // IMPORTANT: even empty string is valid sharedString
    const auto index { string_list_.size() };
    string_list_.append(string);
    string_index_hash_.insert(string, index);
}

bool SharedString::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);
    qsizetype unique_count = 0;
    bool has_unique_count_attr = false;

    while (reader.readNextStartElement()) {
        if (reader.name() == QStringLiteral("sst")) {
            const auto attributes { reader.attributes() };
            if ((has_unique_count_attr = attributes.hasAttribute(QLatin1String("uniqueCount")))) {
                bool ok = false;
                unique_count = attributes.value(QLatin1String("uniqueCount")).toInt(&ok);
                if (!ok) {
                    qDebug("Error: Failed to parse 'uniqueCount' attribute.");
                    return false;
                }
            }
            // Let the loop descend into <sst>'s children
        } else if (reader.name() == QStringLiteral("si")) {
            ParseSharedString(reader);
        } else {
            reader.skipCurrentElement();
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

YXLSX_END_NAMESPACE
