
Tools to access the Red Cross Disaster Transportation Tool programmatically via Python.

This tool automates the generation of daily reports based on data in the DTT and the daily "Avis" reconciliation report.

It has been tested on python 3.8

If you are on Windows: I tested with Windows Subsystem for Linux (WSL), using version 2.  If you have no clue how to do this:
[this is a good page](https://www.windowscentral.com/how-install-wsl2-windows-10).

You will need access to the NHQDCSDLC sharepoint site in order to get acccess to the AVIS reports.

To install and run:

1. install dependencies
    ```shell
    sudo apt install pipenv
    ```
1. check out the repository
    ```shell
    git clone --recursive https://github.com/neilkatin/arc_dtt.git
    cd arc_dtt

    ```
1. use pipenv for dependency resolution (only needed once):
    ```shell
    pipenv --python 3.8
    pipenv install
     ```
1. configure the private environment.  See env.sample for a template.  You'll need types of information:
   1. the DTT URL, username, and password
   1. azure client_id and client_secret -- see the [O365 home page](https://github.com/O365/python-o365#authentication),
   look for "Authentiation Steps".  You will need to use your redcross.org acccount to register the app.
   You will need to add permissions to this application.  The list of permissions is in ```arc_o365/arc_o365.py```.
   Look for the ```_scopes_default``` definition.
1. configure the public (checked into git) environment.  This is in config.py.  You'll almost certainly have to changee
   some values in this file: the DR number and year.
1. The first time you run the program (and after 90 days of non-usage) you will need to initialize your azure oauth 2 tokens.
   A URL will be output on the program output.  Copy and paste that into your browser.  There will be nothing displayed, but
   after the browser has redirected a few times copy the new URL in the browser url bar back into the program window.
   That will initialize your oauth 2 token.
1. run the program
    ```shell
    pipenv run ./main.py --debug
    ```
    This will save a copy of the output in the local folder.  If you add the --store argument then the report will be stored on the server.
    The store location is set in config.py


