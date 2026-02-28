.PHONY: build test publish clean

build:
	dotnet build

test:
	dotnet test

publish:
	dotnet publish src/DocxCli -r osx-arm64 -c Release

clean:
	dotnet clean
	rm -rf publish/
