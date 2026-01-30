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

#ifndef YXLSX_DIMENSION_H
#define YXLSX_DIMENSION_H

#include "namespace.h"
#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

class Dimension {
public:
    Dimension() = default;

    Dimension(int top_row, int left_column, int bottom_row, int right_column);
    explicit Dimension(const QString& dimension);
    explicit Dimension(const char* dimension);

    QString ComposeDimension(bool row_abs = false, bool col_abs = false) const;

    inline bool IsValid() const { return Utility::IsValidCellRange(top_row_, left_column_, bottom_row_, right_column_); }

    inline int TopRow() const { return top_row_; }
    inline int BottomRow() const { return bottom_row_; }
    inline int LeftColumn() const { return left_column_; }
    inline int RightColumn() const { return right_column_; }

    inline void SetTopRow(int row)
    {
        if (top_row_ == kInvalidValue || row < top_row_) {
            top_row_ = row;
        }
    }

    inline void SetBottomRow(int row)
    {
        if (bottom_row_ == kInvalidValue || row > bottom_row_) {
            bottom_row_ = row;
        }
    }

    inline void SetLeftColumn(int col)
    {
        if (left_column_ == kInvalidValue || col < left_column_) {
            left_column_ = col;
        }
    }

    inline void SetRightColumn(int col)
    {
        if (right_column_ == kInvalidValue || col > right_column_) {
            right_column_ = col;
        }
    }

private:
    void Init(const QString& dimension);

private:
    int top_row_ { kInvalidValue };
    int left_column_ { kInvalidValue };
    int bottom_row_ { kInvalidValue };
    int right_column_ { kInvalidValue };
};

QT_END_NAMESPACE_YXLSX

Q_DECLARE_TYPEINFO(YXlsx::Dimension, Q_MOVABLE_TYPE);

#endif // YXLSX_DIMENSION_H
