# OneNote Publish/Exporter

This PowerShell script will export a full hierarchy of all OneNote Notebooks to several supported formats, it will also export all tags such as unresolved todos,
it will also export all attachments.

## Instructions

Specify your settings in the file: /Config/publish.cfg:

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
- Specify if you want to overwrite attachments published or not if existing

Then simply run the script "ExportOneNoteHierarchy.ps1"

## Notes

- the script will not export encrypted objects. TODO is to report this in a warning file
- more Pandoc conversion options TODO
- more MD specific after conversions TODO (did not implement all stuff in the other forks yet)
- Publish on section or notebook level TODO
- Log file instead of Global Error TODO
- Im doubting docx is a good base for md since its does not export the checkboxes while html does
- More metadata would be available for export in md

## Forked

Forked from: https://github.com/SjoerdV/ConvertOneNote2MarkDown and some input from https://github.com/nixsee/ConvertOneNote2MarkDown