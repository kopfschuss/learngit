#!/bin/bash
echo -n -e "hello\tworld\n" "stst" asfb\n; ls && pwd
echo "heool world again!"
echo "clear"
clear

type -a pwd 
clear
type -t if
env
clear
echo $PATH
clear
b=$(ls -l)
echo $b
clear
myvar=USER
b=myvar
echo ${!myvar}
echo $USER
pt=/home/mf/xpdf/xpdf-tools-linux-4.03/bin64/pdftotext
#dir="/mnt/c/Project/workspace/data/pdf"
#file="212_nivo_ipi_SUSAR_bms-2019-034558_10_DE_CA209-901_ROW_open label.pdf"

file="/mnt/c/Project/workspace/data/pdf/212_nivo_ipi_SUSAR_bms-2019-034558_10_DE_CA209-901_ROW_open label.pdf"
$pt
$pt -f 2 "$file" out.txt

