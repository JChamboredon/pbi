pbi is a Python library for accessing and updating PowerBI (.pbi) files.

A typical use would be to download an existing report from the PowerBI service, and automatically alter some text, formatting, or some of the visuals, before publishing the updated report to the service once again.

It can also be used to support different environments: publishing a report still under development to production can hide work-in-progress visuals and replace them with some placeholders showing a default message like 'COMING SOON' or 'WORK IN PROGRESS'.

It can also ensure consistency in the reports by copying titles, logos, and headers from one page to another, or setting specific parameters across the workspace reports (activating the 'keep layer order' option for all visuals, or the search option for all dropdowns, etc.)

More information is available in the pbi wiki.
Browse examples with screenshots to get a quick idea what you can do with pbi.