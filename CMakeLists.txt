cmake_minimum_required(VERSION 3.30)
project(YXlsx VERSION 0.1.0 LANGUAGES CXX)

# ------------------------
# Set Qt6 location manually if not found automatically
# ------------------------
if(APPLE)
    set(Qt6_DIR "/Users/ytx/Qt/6.9.2/macos/lib/cmake/Qt6" CACHE PATH "Qt6 cmake path")
elseif(WIN32)
    set(Qt6_DIR "D:/Qt/6.9.2/mingw_64/lib/cmake/Qt6" CACHE PATH "Qt6 cmake path")
endif()

# ------------------------
# Find required Qt6 modules
# ------------------------
find_package(Qt6 REQUIRED COMPONENTS Gui Core)

# ------------------------
# Standard Qt project setup (Qt 6.9 or higher)
# ------------------------
qt_standard_project_setup(REQUIRES 6.9)

# ------------------------
# Compiler options
# ------------------------
add_compile_options(-Wall -Wextra -Wconversion -Wno-shorten-64-to-32)

set(CMAKE_AUTOMOC ON)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)  # Position Independent Code for shared library

# ------------------------
# Source and header files
# ------------------------
file(GLOB_RECURSE SOURCE_FILES
    "${CMAKE_CURRENT_SOURCE_DIR}/source/*.cc"
)

file(GLOB_RECURSE HEADER_FILES
    "${CMAKE_CURRENT_SOURCE_DIR}/header/api/*.h"
    "${CMAKE_CURRENT_SOURCE_DIR}/header/utils/*.h"
)

# ------------------------
# Build shared library
# ------------------------
add_library(YXlsx SHARED ${SOURCE_FILES} ${HEADER_FILES})
add_library(YXlsx::YXlsx ALIAS YXlsx)

set_target_properties(YXlsx PROPERTIES WINDOWS_EXPORT_ALL_SYMBOLS ON)

target_compile_features(YXlsx PUBLIC cxx_std_20)

# ------------------------
# Define compile-time macros
# ------------------------
target_compile_definitions(YXlsx PRIVATE
    QT_NO_CAST_TO_ASCII
    QT_NO_CAST_FROM_ASCII
    QT_NO_URL_CAST_FROM_STRING
    QT_NO_CAST_FROM_BYTEARRAY
    QT_USE_QSTRINGBUILDER
    QT_USE_FAST_OPERATOR_PLUS
    QT_DISABLE_DEPRECATED_BEFORE=0x060900
)

# ------------------------
# Link Qt libraries
# ------------------------
target_link_libraries(YXlsx
    PUBLIC
        Qt${QT_VERSION_MAJOR}::Core
        Qt${QT_VERSION_MAJOR}::Gui
        Qt${QT_VERSION_MAJOR}::GuiPrivate
)

# ------------------------
# Include directories for consumers
# ------------------------
target_include_directories(YXlsx
    PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/header/api>
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/header/utils>
)

# ------------------------
# Demo / Test target (optional)
# ------------------------
option(BUILD_DEMO "Build YXlsx demo executable" OFF)

if(BUILD_DEMO)
    add_executable(yxlsx_demo main.cpp)

    target_link_libraries(yxlsx_demo
        PRIVATE
            YXlsx
            Qt${QT_VERSION_MAJOR}::Core
            Qt${QT_VERSION_MAJOR}::Gui
            Qt${QT_VERSION_MAJOR}::GuiPrivate
    )
endif()
