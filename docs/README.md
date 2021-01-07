## FORKED!

Forked from: https://github.com/SjoerdV/ConvertOneNote2MarkDown and some input from https://github.com/nixsee/ConvertOneNote2MarkDown so see there for information / download that script.

## WHAT IS IT?

A Script to Export a OneNote Notebook Collection to MD files.

## HOW DOES IT DO THIS?

A PowerShell layer around https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote : The OneNote Application interface includes methods to retrieve, manipulate and update OneNote information and content.

The methods can be broken in four categories:

  - Structure
  - Page Content
  - Navigation
  - Functional

So you can develop a lot of stuff against this by simply calling the Application Interface. (I think the direction is to use the JavaScript API (https://docs.microsoft.com/en-us/graph/integrate-with-onenote) instead).

- The Application Interface publish method out of the box supports exporting to pdf, doc, html, xml, etc... but not md.
- So the OneNote Application publish method is called to create a Word Doc.
- Then the opensource tool Pandoc is used to pick up a published doc and then convert this to md. (In theory this could be done by hand from the XML or the HTML but it is probably easier to do it like that.) (yet creates a dependency)

## STATUS of THIS fork

jan 2021: I forked the repo, did some restructuring, and am in the middle of restructuring. So this fork does not work yet :) (Timeboxed approach). Time for a very first checkin.

- I splitted the onenote calls in modules, so that it, in time is easier to provide more one-note functionality against the modules, against the Application interface, if possible "everything". Also added simple "does it work" test scripts for each module. So they can be tested individually.
- Add a config file since i think that is handier, will also add log module
- Handled some stuff differently
- Probably some more work to get it finished, but have no time anymore :) So will be on TODO :)


