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

#include "coordinate.h"

#include <QDebug>

#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

Coordinate::Coordinate(int row, int column)
    : row_ { row }
    , column_ { column }
{
}

Coordinate::Coordinate(const QString& coordinate) { Init(coordinate); }

Coordinate::Coordinate(const char* coordinate) { Init(QString::fromLatin1(coordinate)); }

void Coordinate::Init(const QString& coordinate)
{
    if (coordinate.isEmpty())
        return;

    // Parse the coordinate string into row and column using Utility::ParseCoordinate
    std::pair<int, int> coord { Utility::ParseCoordinate(coordinate) };

    if (!Utility::CheckCoordinateValid(coord.first, coord.second)) {
        qWarning() << "Invalid coordinate:" << coordinate << "Row:" << coord.first << "Column:" << coord.second;
        return;
    }

    // Assign the parsed row and column values
    row_ = coord.first;
    column_ = coord.second;
}

Coordinate::Coordinate(const Coordinate& other)
    : row_(other.row_)
    , column_(other.column_)
{
}

QT_END_NAMESPACE_YXLSX
