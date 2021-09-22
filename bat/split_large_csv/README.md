split_large_csv

# Index

## split.bat

Dumb way to split a large xlsx file into multiple smaller xlsx files


## Workflow

1. A large xlsx file
2. Save it as a csv file
3. Run the script against the file ( edit the filename inside the script to the name of the target file) , DEFAULT 80000 rows per file
4. Save the result (few csv files ) back to xlsx files
5. Done


###  Supposed to 
```
1. Take a path as parameter/argument of the script to get the target file
2. No need the conversion going back and forth between csv and xlsx file format

```

### What I have done

Hard coded the filename in the script, need to be under the same directory with the target file. Executing the script will split the csv into few smaller csv