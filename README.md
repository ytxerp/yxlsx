# **YXlsx**

- **YXlsx** is a lightweight library for reading and writing Excel (`*.xlsx`) files, designed for seamless integration with [YTX](https://github.com/YtxErp/YTX).
- This library is implemented in **C++** using the **Qt framework**, and is inspired by [QXlsx](https://github.com/QtExcel/QXlsx).

## **Key Features**

- **Efficient and simple:** Provides essential functionality for reading and writing data without handling Excel styles or formatting.
- **Data type support:** Supports reading and writing basic data types, including:
  - Numbers (`int`, `double`, etc.)
  - Booleans (`true/false`)
  - Dates (ISO 8601 format recommended)
  - Shared strings (optimized for Excel's shared string table).

## **Usage Notes**

- **Focus on data:**
    YXlsx is specifically designed as a bridge for handling data in Excel files. It does not process or apply cell formatting, styling, or advanced features.
- **Data validation:**
    Ensure that your data is validated and preprocessed before passing it to the library for optimal results.

## **Installation**

- Clone the [yxlsx](https://github.com/ytxerp/yxlsx) repository into the `external/` folder of the project:

      ```bash
      git clone https://github.com/ytxerp/yxlsx.git external/yxlsx
      ```

- Add it in CMakeLists:

      ```cmake
      add_subdirectory(external/yxlsx/yxlsx/)
      ```

## **Acknowledgments**

- A special thanks to the [QXlsx](https://github.com/QtExcel/QXlsx) project for providing the foundation on which this library was built.

## **License**

- YXlsx is released under the [MIT License](LICENSE). Feel free to use, modify, and distribute this library in your projects.
