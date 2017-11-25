ROOT_DIR := $(patsubst %/,%, $(dir $(lastword $(MAKEFILE_LIST))))

ROOT_SOURCE_DIR := $(ROOT_DIR)/src

ROOT_OUTPUT_DIR := $(ROOT_DIR)/out

main.ps1 := $(ROOT_OUTPUT_DIR)/main.ps1


mk_target_dir = $(if $(wildcard ${@D}),,mkdir -p ${@D} &&)

all:

clean: ; rm -rf $(ROOT_OUTPUT_DIR)

all : $(main.ps1)

#
#  Merge .cs files reordering 'using ...' statements.
#
main.ps1: $(main.ps1)
$(main.ps1):    $(ROOT_SOURCE_DIR)/MsiTools.cs \
                $(ROOT_SOURCE_DIR)/Win32Interop.cs \
                $(ROOT_SOURCE_DIR)/MsiTools.ps1 \
;   $(mk_target_dir) \
    { \
        echo '$$Source = @"' && \
        cat $(filter %.cs, $^) | awk '  /using [^;]*;/ { print }' | sort -u && \
        cat $(filter %.cs, $^) | awk '! /using [^;]*;/ { print }' && \
        echo '"@' && \
        echo 'Add-Type -TypeDefinition $$Source -Language CSharp' && \
        cat $(filter %.ps1, $^); \
    } | unix2dos > $@
