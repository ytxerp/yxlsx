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

#include "zipreader.h"

QT_BEGIN_NAMESPACE_YXLSX

ZipReader::ZipReader(const QString& file_path)
    : reader_ { new QZipReader(file_path) }
{
    Init();
}

ZipReader::ZipReader(QIODevice* device)
    : reader_ { new QZipReader(device) }
{
    Init();
}

void ZipReader::Init()
{
    file_path_.clear();

    for (const auto& file : reader_->fileInfoList()) {
        if (file.isFile || (!file.isDir && !file.isSymLink)) {
            file_path_.append(file.filePath);
        }
    }
}

QT_END_NAMESPACE_YXLSX
