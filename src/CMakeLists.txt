cmake_minimum_required(VERSION 3.6 FATAL_ERROR)

if (WIN32)
    ENABLE_LANGUAGE(RC)
endif()

include(FetchContent)

FetchContent_Declare(
    simstr
    GIT_REPOSITORY https://github.com/orefkov/simstr.git
    GIT_SHALLOW TRUE
    GIT_TAG tags/rel1.3.1
    FIND_PACKAGE_ARGS NAMES simstr
)
set(SIMSTR_BUILD_TESTS OFF)
set(SIMSTR_BENCHMARKS OFF)
FetchContent_MakeAvailable(simstr)

FetchContent_Declare(
    openxlsx
    GIT_REPOSITORY https://github.com/troldal/OpenXLSX.git
    GIT_SHALLOW TRUE
    GIT_TAG v0.5.1
)
set(BUILD_TESTS OFF)
FetchContent_MakeAvailable(openxlsx)

FetchContent_Declare(
    libxls
    GIT_REPOSITORY https://github.com/libxls/libxls.git
    GIT_SHALLOW TRUE
    GIT_TAG v1.6.3
)
FetchContent_MakeAvailable(libxls)

set(LIBXLS_SRC_DIR "${CMAKE_BINARY_DIR}/_deps/libxls-src/src")
set(LIBXLS_INC_DIR "${CMAKE_BINARY_DIR}/_deps/libxls-src/include")

file(WRITE "${LIBXLS_SRC_DIR}/config.h"
"/* config.h for libxls on Windows with MSVC */\n"
"#define HAVE_LOCALE_H 1\n"
"#define HAVE_STDLIB_H 1\n"
"#define HAVE_STRING_H 1\n"
"#define HAVE_STRINGS_H 0\n"
"#define HAVE_SYS_STAT_H 1\n"
"#define HAVE_SYS_TYPES_H 1\n"
"#define HAVE_UNISTD_H 0\n"
"#define HAVE_WCHAR_H 1\n"
"#define STDC_HEADERS 1\n"
"#define PACKAGE_VERSION \"1.6.3\"\n"
)

add_library(libxls STATIC
    ${LIBXLS_SRC_DIR}/endian.c
    ${LIBXLS_SRC_DIR}/locale.c
    ${LIBXLS_SRC_DIR}/ole.c
    ${LIBXLS_SRC_DIR}/xls.c
    ${LIBXLS_SRC_DIR}/xlstool.c
)
target_include_directories(libxls PUBLIC ${LIBXLS_INC_DIR} ${LIBXLS_SRC_DIR})
target_compile_definitions(libxls PRIVATE _CRT_SECURE_NO_WARNINGS PACKAGE_VERSION="1.6.3" strdup=_strdup)
set_target_properties(libxls PROPERTIES C_STANDARD 99)

set(XLS21C_SRC
    xls21c_addin.cpp
    xls21c_addin.h
    xlsx_reader.cpp
    xlsx_reader.h
    xls_helper.c
    xls_helper.h
    stdafx.cpp
    stdafx.h
    addin_base.cpp
)

if (WIN32)
    list(APPEND XLS21C_SRC dllmain.cpp xls21c.rc xls21c.def)
endif()

add_library(${PROJECT_NAME} SHARED ${XLS21C_SRC})
target_link_libraries(${PROJECT_NAME} simstr::simstr OpenXLSX libxls)

if (WIN32)
    add_compile_definitions(
        UNICODE
        _WINDOWS
        _USRDLL
        _CRT_SECURE_NO_WARNINGS
        _CRT_SECURE_NO_DEPRECATE
        ADDINCPP_EXPORTS
    )
    if (GCC_COMPILER)
        target_compile_options(${PROJECT_NAME} PRIVATE -fpermissive)
    endif()
else()
    if (XLS21C_STATIC_RUNTIME)
        target_link_options(${PROJECT_NAME} PRIVATE -static-libgcc -static-libstdc++)
    endif()
    target_link_options(${PROJECT_NAME} PRIVATE -Wl,--no-undefined -Xlinker --version-script -Xlinker ${CMAKE_CURRENT_SOURCE_DIR}/version.script)
    target_link_libraries(${PROJECT_NAME} pthread)
endif()

set_target_properties(${PROJECT_NAME} PROPERTIES POSITION_INDEPENDENT_CODE ON)

if (NOT WIN32)
    add_custom_command(
        TARGET ${PROJECT_NAME} POST_BUILD
        DEPENDS ${PROJECT_NAME}
        COMMAND $<$<CONFIG:release>:strip>
        ARGS --strip-all $<TARGET_FILE:${PROJECT_NAME}>
    )
endif()
