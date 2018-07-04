# XSLX Row Remover

This is an example project to remove rows from an XLSX file.

## Basic Usage
```
./bin/<machine architecture>/xslxrowremover -i <inputXLSXfilename> -o <outputXLSXfilename> -r <range>
```

### Prerequisites

To run the binary only shell access is required

### Linux Example
```
./bin/linux_amd64/xslxrowremover -i Book1.xlsx -o Book1New.xlsx -r 2:4
```
Above command will parse Book1.xlsx and remove rows from index 2 to 4 (index starts from 0)

### Linux Example 2
```
./bin/linux_amd64/xslxrowremover -i Book1.xlsx -o Book1New.xlsx -r 2
```
Above command will parse Book1.xlsx and remove row of index 2 

## Compiling


You need go lang installed in you machine to compile the source. Refer [Go Programing Language, getting started guide](https://golang.org/doc/install)


```
go get "github.com/360EntSecGroup-Skylar/excelize"
go build
```

To build for a specific target OS and architecture 
```
GOOS=solaris GOARCH=amd64 go build
```
There is a good [guide by Digital Ocean](https://www.digitalocean.com/community/tutorials/how-to-build-go-executables-for-multiple-platforms-on-ubuntu-16-04) for cross-compiling fro different platforms


## Credits

* Used Excelize library [360EntSecGroup-Skylar/excelize](https://github.com/360EntSecGroup-Skylar/excelize)
* Inspired by [tealeg/xlsx](https://github.com/tealeg/xlsx)

