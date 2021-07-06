
Tools to access the Red Cross Disaster Transportation Tool programmatically via Python.

This tool automates the generation of daily reports based on data in the DTT and the daily "Avis" reconciliation report.

It has been tested on python 3.8

To install and run:

1. check out the repository
1. use pipenv for dependency resolution:
     ```shell
     pipenv --python 3.8
     pipenv install
     ```
1. run the program
     ```shell
     pipenv run ./main --debug
     ```
     This will save a copy of the output in the local folder.  If you add the --store argument then the report will be stored on the server.
     The store location is set in config.py


