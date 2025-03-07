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

#include "dimension.h"

#include <QDebug>
#include <QList>

#include "utility.h"

QT_BEGIN_NAMESPACE_YXLSX

Dimension::Dimension(int top_row, int left_column, int bottom_row, int right_column)
    : top_row_ { top_row }
    , left_column_ { left_column }
    , bottom_row_ { bottom_row }
    , right_column_ { right_column }
{
}

Dimension::Dimension(const QString& dimension) { Init(dimension); }

Dimension::Dimension(const char* dimension) { Init(QString::fromLatin1(dimension)); }

void Dimension::Init(const QString& dimension)
{
    // Split the dimension string by ':' into parts.
    const QStringList parts { dimension.split(QLatin1Char(':'), Qt::SkipEmptyParts) };

    // Check if the dimension string is empty or contains more than two parts (invalid format).
    if (parts.isEmpty() || parts.size() >= 3) {
        qWarning() << "Invalid dimension string:" << dimension;
        return;
    }

    // Parse the start and end coordinates.
    // If there is only one part, the start and end coordinates will be the same.
    std::pair<int, int> start { Utility::ParseCoordinate(parts.first()) };
    std::pair<int, int> end { Utility::ParseCoordinate(parts.size() == 2 ? parts.last() : parts.first()) };

    if (!Utility::CheckDimensionValid(start.first, start.second, end.first, end.second)) {
        qWarning() << "Invalid dimension coordinates." << dimension;
        return;
    }

    // Initialize the row and column range.
    top_row_ = start.first;
    left_column_ = start.second;
    bottom_row_ = end.first;
    right_column_ = end.second;
}

QString Dimension::ComposeDimension(bool row_abs, bool col_abs) const
{
    if (!CheckValid()) {
        return {};
    }

    const QString top { Utility::ComposeCoordinate(top_row_, left_column_, row_abs, col_abs) };

    if (top_row_ == bottom_row_ && left_column_ == right_column_) {
        return top;
    }

    const QString bottom { Utility::ComposeCoordinate(bottom_row_, right_column_, row_abs, col_abs) };
    return QStringLiteral("%1:%2").arg(top, bottom);
}

QT_END_NAMESPACE_YXLSX
