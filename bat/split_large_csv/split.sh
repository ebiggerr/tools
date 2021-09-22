#!/bin/bash
FILENAME=test.csv
HDR=$(head -1 $FILENAME)
split -l 80000 $FILENAME xyz
n=1
for f in xyz*
do
     if [ $n -gt 1 ]; then 
          echo $HDR > Part${n}-$FILENAME
     fi
     cat $f >> Part${n}-$FILENAME
     rm $f
     ((n++))
done