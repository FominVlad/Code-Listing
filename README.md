# Create docx report file from sources

## How to build

```cmd
dotnet publish -r win-x64 -c Release -o out /p:PublishSingleFile=true
```

## How to run

```
Files2Doc.exe -f ZemResControls -o Reports/ZemResControls.docx
```



## Usage:

```
>Files2Doc.exe --help
Files2Doc 1.0.0
Copyright (C) 2020 Files2Doc

  -v, --verbose                Set output to verbose messages.

  -x, --extensions             (Default: *.cs *.aspx *.ascx *.ashx *.js *.html *.resx *.csproj *.xsd *.xss *.xss)
                               Extensions to process

  -f, --folders                (Default: .) Folders to process

  -r, --remove-comments        (Default: true) Remove comments from source files

  -m, --remove-xml-comments    (Default: true) Remove comments from XML files

  -o, --output                 (Default: output.docx) Output file name

  --help                       Display this help screen.

  --version                    Display version information.

```

