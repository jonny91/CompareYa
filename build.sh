#/bin/bash

# mac
CGO_ENABLED=0
GOOS=darwin
GOARCH=amd64
go build -o ./CompareYa .
chmod a+x ./CompareYa

#windows
CGO_ENABLED=0
GOOS=windows
GOARCH=amd64
go build -o ./CompareYa.exe .