# OneNote Publish/Exporter

This PowerShell script will export a full hierarchy of all OneNote Notebooks to several supported formats on page level.

## Instructions

Edit Export/Config/export.cfg:

- Specify one or more export formats desired:

   - one (OneNote files)
   - mht (mhtml files, web page archive format)
   - pdf  (PDF files)
   - xps  (open XML Paper specification)
   - docx (Microsoft Word in newer .docx format)
   - doc  (Microsoft Word classic .doc format)
   - emf (Enhance Metafile)
   - htm (HTML files, with attachments in same location)
   - markdown (Pandocs markdown)
   - markdown_mmd (MultiMarkdown)
   - markdown_phpextra (PHP Markdown extra)
   - markdown_strict (original unextended Markdown)
   - commonmark (CommonMark Markdown)
   - commonmark_x (commonMark Markdown with Extensions)
   - gfm (Github Flavored Markdown)
   - markdown_github (use only markdown github if you need extensions not supported in gfm)

- Specify the folder you would like to export to
- Specify if you want to overwrite the OneNote Published file or not if existing

Run the script "ExportOneNoteHierarchy.ps1"

The script will autodownload Pandoc for conversion when needed unless stated in the config file this is not desired.

## Notes

- the script will not export encrypted objects. TODO is to report this in a warning file
- the script will not unfold folded pages, this is TODO
- more Pandoc conversion options TODO
- more MD specific after conversions TODO I did not implement all in the other forks yet
- Publish on section or notebook level TODO
- Log file instead of Global Error TODO

## Forked

Forked from: https://github.com/SjoerdV/ConvertOneNote2MarkDown and some input from https://github.com/nixsee/ConvertOneNote2MarkDown