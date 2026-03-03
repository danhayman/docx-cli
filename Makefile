.PHONY: build test publish clean

build:
	dotnet build src/ox.slnx

test:
	dotnet test src/ox.slnx

publish:
	dotnet publish src/Ox -r osx-arm64 -c Release

clean:
	dotnet clean src/ox.slnx
	rm -rf publish/
