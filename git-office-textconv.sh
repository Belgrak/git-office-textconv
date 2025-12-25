#!/bin/bash

FILEPATH="$1"

# Проверка существования файла
if [ ! -f "$FILEPATH" ]; then
    echo "Error: File not found."
    exit 1
fi

# Расширение файла в нижнем регистре
EXT="${FILEPATH##*.}"
EXT="${EXT,,}"

# Проверка наличия команды
has_cmd() {
    command -v "$1" >/dev/null 2>&1
}

# Преобразуем файлы в текст
case "$EXT" in
    pdf)
        if has_cmd pdftotext; then
            pdftotext -layout -nopgbrk "$FILEPATH" -
        else
            echo "[WARN] pdftotext not found. Install poppler-utils"
        fi
        ;;
    docx|odt|epub)
        if has_cmd pandoc; then
            pandoc "$FILEPATH" -t markdown
        elif has_cmd docx2txt && [ "$EXT" == "docx" ]; then
            docx2txt "$FILEPATH" -
        elif has_cmd odt2txt && [ "$EXT" == "odt" ]; then
            odt2txt "$FILEPATH"
        else
            echo "[WARN] pandoc not found."
        fi
        ;;

    doc)
        if has_cmd antiword; then
            antiword "$FILEPATH" 2>/dev/null || strings "$FILEPATH"
        elif has_cmd catdoc; then
            catdoc "$FILEPATH" 2>/dev/null || strings "$FILEPATH"
        else
            echo "[WARN] antiword or catdoc not found."
        fi
        ;;

    xlsx|xls|ods)
        if has_cmd ssconvert; then
            ssconvert --export-type=Gnumeric_stf:stf_csv "$FILEPATH" fd://1
        elif has_cmd xlsx2csv && [ "$EXT" == "xlsx" ]; then
            xlsx2csv "$FILEPATH"
        elif has_cmd xls2csv && [ "$EXT" == "xls" ]; then
            xls2csv "$FILEPATH"
        else
            echo "[WARN] ssconvert (gnumeric) or xlsx2csv not found."
        fi
        ;;

    ppt|pptx|odp)
        if has_cmd pptx2md && [[ "$EXT" == "pptx" ]]; then
             pptx2md "$FILEPATH"
        elif has_cmd catppt && [[ "$EXT" == "ppt" ]]; then
            catppt "$FILEPATH"
        elif has_cmd soffice; then
             tmp_dir=$(mktemp -d)
             soffice --headless --convert-to txt:Text --outdir "$tmp_dir" "$FILEPATH" >/dev/null 2>&1
             cat "$tmp_dir"/*.txt
             rm -rf "$tmp_dir"
        else
            echo "[WARN] No suitable tool for slides found."
        fi
        ;;

    *)
        echo "[INFO] Unknown format, extracting strings..."
        strings "$FILEPATH"
        ;;
esac