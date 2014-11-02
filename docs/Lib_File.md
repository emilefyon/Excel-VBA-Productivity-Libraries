# File Library

* Name file: LIB_File.bas

## Function lists

* [writeFile](#writefile-file-content)
* [readFile](#readfile-file)
* [readfileAndTruncate](#readfileandtruncate-file)
* [fileExists](#fileexists-file)

### writeFile (file, content)

Overwrite the file specified with the content specified

#### Arguments
* file as String: the full path of the string
* content as String: the content that has to be written in the file

#### Specifications / limitations
* If the file does not exists, the file is created
* The folder has to exists

*Example: writeFile("c:\NewFile.txt", "Lorem Ipsum") *




### readFile (file)

Read the content of a file and return a single line with all the content

#### Arguments
* file as String : the full path of the file
* The content is retrieved without any line returns (line returns are replaced by space)

#### Specifications / limitations
* The file has to Exists, currently no error handling


### readfileAndTruncate (file)

calls readFile() and then truncate the text to 30000 characters in order to avoid Excel limitations in cell content

#### Arguments
* file as String : the full path of the file


#### Specifications / limitations
* The file has to Exists, no error handling
* The content is retrieved without any line returns (line returns are replaced by space)
* Only the 30.000 first characters are retrieved

### fileExists (file)

Check if the specified file exists. Return a boolean value.

#### Arguments
* file as String : the full path of the file


#### Specifications / limitations
* none






