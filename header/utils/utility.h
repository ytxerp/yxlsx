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

YXLSX_BEGIN_NAMESPACE

struct CellAddress {
    int row { -1 };
    int column { -1 };
    bool valid { false };

    bool IsValid() const { return valid; }
    bool IsBeforeOrEqual(const CellAddress& other) const { return row <= other.row && column <= other.column; }
};

class Utility {
public:
    static CellAddress ParseCoordinate(const QString& coordinate);
    static QString ComposeCoordinate(int row, int column, bool row_abs = false, bool col_abs = false);
    static QStringList SplitPath(const QString& path);
    static QString GetRelFilePath(const QString& filePath);

    static QString GenerateSheetName(const QStringList& sheet_names, const QString& name_proposal, int& last_sheet_index);
    static QString UnescapeSheetName(const QString& sheetName);

    static bool IsSpacePreserveNeeded(const QString& string);
    static constexpr bool IsValidRowColumn(int row, int column) { return row >= 1 && row <= kMaxExcelRow && column >= 1 && column <= kMaxExcelColumn; }

private:
    static QString ComposeColumn(int column);
};

YXLSX_END_NAMESPACE
#endif // YXLSX_UTILITY_H
