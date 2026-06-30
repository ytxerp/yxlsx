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

YXLSX_BEGIN_NAMESPACE

class Dimension final {
public:
    Dimension() = default;
    explicit Dimension(const QString& dimension);

    QString ComposeDimension(bool row_abs = false, bool col_abs = false) const;

    inline bool IsValid() const
    {
        return top_row_ != kInvalidValue && left_column_ != kInvalidValue && bottom_row_ != kInvalidValue && right_column_ != kInvalidValue && top_row_ >= 1
            && bottom_row_ >= top_row_ && left_column_ >= 1 && right_column_ >= left_column_ && bottom_row_ <= kMaxExcelRow && right_column_ <= kMaxExcelColumn;
    }

    inline void Extend(int row, int col)
    {
        // -------------------------
        // Row (top / bottom)
        // -------------------------
        top_row_ = (top_row_ == kInvalidValue) ? row : std::min(top_row_, row);
        bottom_row_ = (bottom_row_ == kInvalidValue) ? row : std::max(bottom_row_, row);

        // -------------------------
        // Column (left / right)
        // -------------------------
        left_column_ = (left_column_ == kInvalidValue) ? col : std::min(left_column_, col);
        right_column_ = (right_column_ == kInvalidValue) ? col : std::max(right_column_, col);
    }

private:
    int top_row_ { kInvalidValue };
    int left_column_ { kInvalidValue };
    int bottom_row_ { kInvalidValue };
    int right_column_ { kInvalidValue };
};

YXLSX_END_NAMESPACE

Q_DECLARE_TYPEINFO(yxlsx::Dimension, Q_MOVABLE_TYPE);

#endif // YXLSX_DIMENSION_H
