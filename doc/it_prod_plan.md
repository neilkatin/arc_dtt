Product Spec for DTT Reports
============================

Background
----------

We currently have a sample set of reports we are using to help the Transportation function.
This document documents one of those reports: all in-person responders are sent an email
with the vehicles in their name (or a message that they have no vehicles).

Usage has shown that the report is valuable in discovering discrepancies between the Disaster
Transportation Tool (DTT) database, and the on-the-ground reality.

Document Organization
---------------------

This document is divided into parts:
* documenting the reports themselves
* specifying configuration information needed to control the delivery of the reports
* features we don't currently do but would like if the implementation isn't too difficult

The vehicle emails
------------------

Everyone with a vehicle assigned in the DTT gets an email showing some basic background information,
and information about each vehicle in their name.

The subject of the email is template driven:

```jinja
Vehicle Status - {{ date }} - {{ first_name }} {{ last_name }}
```

Here is the template used to send the email:

```jinja
<html>
  <head>
      <title>Vehicle report for  {{ first_name }} {{ last_name }}</title>
  </head>
  <body>

    <p>
    This is an automated report to try to more accurately track our vehicles.
    </p>

    <p>
    This report was sent to {{ first_name }} {{ last_name }}.
    We think you have the following vehicle{{ vehicles|length|pluralize }}:
    </p>

    {% for row in vehicles %}
        {% set v = row['Vehicle'] %}
        {% set code = v['VehicleCategoryCode'] %}
        {% set code_string = vehicle_category_code_to_string(code) %}
        <ul>
            {% if code_string is not none %}
            <li>Vehicle category: {{ code_string }}</li>
            {% endif %}
            {% if v.Vendor is not none %}
            <li>Vendor: {{ v.Vendor }}</li>
            {% endif %}
            <li>Vehicle id: {{ v.KeyNumber }}
            {% if v.Vendor == 'Avis' %}
                (called "MVA" on the key tag)
            {% endif %}
            </li>
            <li>License plate: {{ v.PlateState }} {{ v.Plate }}
        </ul>
    {% endfor %}

    <p>
    If all that is correct: YOU DON'T NEED TO DO ANYTHING.  If its wrong: please email
    <a href='mailto:{{ reply_email }}'>{{ reply_email }}</a>.
    </p>

    <p>
    This email generated on {{ date }}
    </p>
  </body>
</html>
```

(the template uses the [Jinja2](https://jinja.palletsprojects.com/en/3.0.x/) template language)

The report contains some basic boiler-plate information:
* why the recipient is getting the report
* who the report was sent to
* when the report was generated
* who to send comments to (which is also the from address)

Then for each vehicle it shows:
* what type of vehicle it is (taken directly from the DTT "Ctg" column in the "Vehicles" tab)
  * examples: ERV, Box Truck, Chapter Vehicle, Rental, etc
* The vendor if there is one
  * example: Penske, Red Cross, Avis
* The vehicle ID (aka MVA number)
* The vehicle license plate state and number

Multiple vehicles assigned to the same person are listed sequentially

The no-vehicle emails
---------------------

We also send an email to everyone who we don't think has a vehicle.
The logic we use for determinining this set of people is:
* people assigned or checked into the DR (as shown on the Staff Roster)
* who are not virtual
* who did not have a vehicle assigned

The subject for this email is template driven:

```jinja
Vehicle Status - {{ date }} - {{ first_name }} {{ last_name }}
```

The template for this email is:

```jinja
<html>
  <head>
      <title>Vehicle report for  {{ first_name }} {{ last_name }}</title>
  </head>
  <body>

    <p>
    This is an automated report to try to more accurately track our vehicles.
    </p>

    <p>
    We think you do not currently have a Red Cross vehicle on the DRO.
    </p>

    <p>
    If you don't have a Red Cross vehicle: YOU DON'T NEED TO DO ANYTHING.
    </p>

    <p>
    If you do have a Red Cross vehicle (chapter or rental): please reply to this message with your vehicle info so we can update our database.
    </p>

    <p>
    Thanks.  This email generated on {{ date }}
    </p>
  </body>
</html>
```

Configuration of the system
---------------------------

We needed to configuration the following information, on a per-DR basis:
* Sending Address
  * also gets bounced messages on errors
* Auto-sending
  * what time to send
  * often send (currently: every X days)
  * separately controllable for "have vehicle" and "no vehicle" reports
  * ability to disable the running of the report for both have and no vehicle reports.
* Manual sending on demand
* Test sending
  * to an alternate test email instead of the 'real' email
  * (to confirm the configuration is correct; may not be necessary)
* a log of the batch run
  * timestamp
  * who the messages were sent to
  * what was information sent

Since DROs vary widely, we would like to have the following controls on a DR basis:
* customization of the email subject and body templates, initialzed from a common default
* which vehicle categories to report on (based on the DTT Category)

Additional features
-------------------

The current reports assume they are emails because we didn't have access to sending SMS messages.
But the DTT already sends text messages; using texts instead of emails would probably be friendlier
to our responders.

The current reports say to send an email if there is information that needs changing.  But
a more positive acknowledgement would be nice:
* a link to an everythings OK page
* a link to a needs-updating page to collect update information
* a report of responses
  * who didn't respond
  * outstanding updates requested
  * a small bit of workflow support
    * who processed update
    * when it was updated  
