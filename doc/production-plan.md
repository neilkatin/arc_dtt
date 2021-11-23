
Getting reporting tools into production
=======================================

Draft as of 2021-10-20
Neil Katin

Document goals
-------------

The goal for this document is to give an overview of the transportation reporting application, and to start a dialog about how
to better 'productize' the application.  The goal is to provide background for starting a discussion and not to insist on
a particular implementation plan.

Application Overview
--------

We have a set of transportation reporting tools that automate manual tasks that the transportation lead is supposed to do.
The tools currently do three main tasks:

1. Reconcile the current Avis vehicles listed in the DTT against the Avis report to show discrepancies,
   missing vehicles, and vehicles have been returned but not released in the DTT
1. Prepare and format the group vehicle report (which shows all vehicles on teh DRO).  The DTT has a version of this,
   but the DTT report only includes rental vehicles and takes manual editing of the report every day.
1. Messages to vehicle holders and others about the DTTs

The tools have been used to support 4 DRs so far; they have been valuable enough that we would like to
figure out how to deploy them more widely, in a supportable way.

Deployment Environment
------------------------------

The tools currently are running on a personal linux server.  This is obviously not ideal.
We would like to have it run in a more reliable production environment and be able to
be supported by a wider set of people.

The tools are implemented in the Python language, and have no UI.  They are run from a schedule (linux 'crontab') on a periodic basis.

We would like hosted in the cloud.  A proposal that makes sense to us is:

* use Azure Functions as the cloud implementation.  Most of the time the tools do not need to be resident in memory;
  Azure Functions only charges for the time the program is actually running.
* use Timers to configure when the tools run, and for which DR
* configure the timers from files in sharepoint, to allow easy adding of new DRs and control.

The application takes about 15 seconds to run for each DRO and report.  The total hosting cost should be on the order
of tens of dollars per month.


Current implementation
----------------------

The current implementation is available on GitHub ([neilkatin/arc_dtt](https://github.com/neilkatin/arc_dtt)).
Dependencies and runtime isolation is managed via [pipenv](https://github.com/pypa/pipenv).  There is currently
no UI or remote access, so there are limited security perimeter.

The primary APIs accessed are Microsoft Graph for access to email and sharepoint.  It also accesses the DTT's API to
obtain vehicle and responder information.

The current implementation uses the developer's and the generic DTT DRnnn-nnLog-TraN@redcross.org accounts for accessing data.
It would be much better if we could get a generic account to access microsoft graph and the DTT instead.


Glossary
--------

* DTT - Disaster Transportation Tool, the tool we use to track vehicles assigned to a DR
