## Test environments
* Local OS X (darwin15.6.0) install, R version 3.4.2
* Ubuntu 14.04.5 LTS (on travis-ci), R version 3.4.2
* Local Windows 10 install, R 3.4.2, and win_builder

## R CMD check results
There were no ERRORs, WARNINGs or NOTEs.

## Steps taken since the previous version:
- Package would not work correctly when installed from the pre-compiled files due to issues with lazy loading. Error fixed and package now works as expected. See issue 108 on GitHub.

## Steps taken to address previous CRAN comments:

- Removed the redundant 'from R' from package description
- Added an example for every user-accessible function to the .md files
- I've made sure that none of the examples write to the user's home directory.  

## Downstream dependencies
There are currently no downstream dependencies for this package
