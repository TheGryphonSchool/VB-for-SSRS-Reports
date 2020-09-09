# VB-for-SSRS-Reports
This repo contains a folder named for each SSRS report that needs custom code,
and a folder named `utilities` where repurposable code lives.

Although the code is organised in seperate, modular files, each report has one
file which is the concatenates all the necessary source files. This file is
named according to this convention: `Project/_Project.vb`.

## Contributing
Although you can work on the project in any environment, I recommend using
VScode or even Visual Studio, to make use of the Combine Files extension, which
is configured for the project and used to concatenate the source files for each
report into the output file.

Once you install the Combine Files extension, you can use it by right-clicking
the parent folder or any blank space in VS(code)'s Explorer window, and click
`Combine Files`.
The config settings for the Combine Files extension are in
`.combinefilesrc.json`.

All source files in the report folder are automatically included in the output
(unless they begin with a `_`). To include a utility file, add a glob for it.
