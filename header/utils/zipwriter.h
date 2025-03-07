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

#ifndef YXLSX_ZIPWRITER_H
#define YXLSX_ZIPWRITER_H

#include <private/qzipwriter_p.h>

#include <QIODevice>
#include <QScopedPointer>

#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

class ZipWriter {
    Q_DISABLE_COPY(ZipWriter)

public:
    explicit ZipWriter(const QString& file_path);
    explicit ZipWriter(QIODevice* device);
    ~ZipWriter() = default;

    inline void AddFile(const QString& file_path, QIODevice* device) { writer_->addFile(file_path, device); }
    inline void AddFile(const QString& file_path, const QByteArray& data) { writer_->addFile(file_path, data); }

    inline bool IsError() const { return writer_->status() != QZipWriter::NoError; }
    inline void Close() { writer_->close(); }

private:
    QScopedPointer<QZipWriter> writer_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_ZIPWRITER_H
