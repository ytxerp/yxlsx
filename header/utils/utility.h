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

#ifndef YXLSX_UTILITY_H
#define YXLSX_UTILITY_H

#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

class Utility {
public:
    static std::pair<int, int> ParseCoordinate(const QString& coordinate);
    static QString ComposeCoordinate(int row, int column, bool row_abs = false, bool col_abs = false);
    static QStringList SplitPath(const QString& path);
    static QString GetRelFilePath(const QString& filePath);

    static QString GenerateSheetName(const QStringList& sheet_names, const QString& name_proposal, int& last_sheet_index);
    static QString UnescapeSheetName(const QString& sheetName);

    static bool IsSpaceReserveNeeded(const QString& string);

    static constexpr bool IsValidRowColumn(int row, int column) { return row >= 1 && row <= kExcelRowMax && column >= 1 && column <= kExcelColumnMax; }
    static bool IsValidCellRange(int top_row, int left_column, int bottom_row, int right_column)
    {
        return Utility::IsValidRowColumn(top_row, left_column) && Utility::IsValidRowColumn(bottom_row, right_column) && top_row <= bottom_row
            && left_column <= right_column;
    }

private:
    static int ParseColumn(const QString& column);
    static QString ComposeColumn(int column);
};

QT_END_NAMESPACE_YXLSX
#endif // YXLSX_UTILITY_H
