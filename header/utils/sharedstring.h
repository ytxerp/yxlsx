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

#ifndef YXLSX_SHAREDSTRING_H
#define YXLSX_SHAREDSTRING_H

#include <QHash>
#include <QIODevice>
#include <QList>
#include <QSet>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "abstractooxmlfile.h"
#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

class SharedString final : public AbstractOOXmlFile {
public:
    explicit SharedString(OperationMode mode);

    int SetSharedString(const QString& string, int row, int column);
    void IncrementReference(int index, int row, int column);

    QSet<std::pair<int, int>> RemoveSharedString(const QString& string, int row, int column);

    inline int GetSharedStringIndex(const QString& string) const { return string_index_hash_.value(string, -1); }
    inline const QList<QString>& GetSharedString() const { return string_list_; }
    inline bool IsEmpty() const { return string_list_.isEmpty(); }

    QString GetSharedString(int index) const;

    void ComposeXml(QIODevice* device) const override;
    bool ParseXml(QIODevice* device) override;

private:
    void ParseSharedString(QXmlStreamReader& reader); // <si>

private:
    /**
     * @brief Simulates the shared strings in the filename.xlsx/xl/sharedStrings.xml file.
     *
     * @details This list determines the order and indices of shared strings in the table.
     * Each string is stored in sequence, and the index corresponds to its position.
     *
     * @note This structure is critical for XML-based Excel file parsing.
     */
    QList<QString> string_list_ {};

    /**
     * @brief string_index_hash_
     *
     * Maintains a mapping between shared strings and their corresponding indices
     * in the shared string table (`string_list_`). This allows efficient lookup
     * of a string's index when inserting or referencing shared strings.
     */
    QHash<QString, int> string_index_hash_ {};

    /**
     * @brief string_coordinate_hash_
     *
     * Maintains a mapping between shared strings and the cell coordinates
     * where they are used. Each string is associated with a nested hash table,
     * where the keys represent row numbers and the values represent column numbers.
     * @note The index of a shared string can change over time as new strings are added or removed.
     *       Therefore, using the index as a key is not recommended.
     */
    QHash<QString, QSet<std::pair<int, int>>> string_coordinate_hash_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_SHAREDSTRING_H
