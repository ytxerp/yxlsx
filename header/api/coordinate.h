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

#ifndef YXLSX_COORDINATE_H
#define YXLSX_COORDINATE_H

#include "namespace.h"
#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

class Coordinate {
public:
    Coordinate() = default;
    ~Coordinate() = default;

    Coordinate(int row, int column);
    Coordinate(const QString& coordinate);
    Coordinate(const char* coordinate);
    Coordinate(const Coordinate& other);

    inline void SetRow(int row) { row_ = row; }
    inline void SetColumn(int col) { column_ = col; }
    inline int Row() const { return row_; }
    inline int Column() const { return column_; }

    inline bool operator==(const Coordinate& other) const { return row_ == other.row_ && column_ == other.column_; }
    inline bool operator!=(const Coordinate& other) const { return row_ != other.row_ || column_ != other.column_; }

    static constexpr bool CheckValid(const Coordinate& coordinate) { return Utility::CheckCoordinateValid(coordinate.Row(), coordinate.Column()); }

private:
    void Init(const QString& coordinate);

private:
    int row_ { kInvalidValue };
    int column_ { kInvalidValue };
};

QT_END_NAMESPACE_YXLSX

Q_DECLARE_TYPEINFO(YXlsx::Coordinate, Q_MOVABLE_TYPE);

#endif // YXLSX_COORDINATE_H
