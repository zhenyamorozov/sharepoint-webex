=====================================
Sharepoint Lists and Webex Automation
=====================================
.. image:: https://static.production.devnetcloud.com/codeexchange/assets/images/devnet-published.svg
    :alt: published
    :target: https://developer.cisco.com/codeexchange/github/repo/zhenyamorozov/sharepoint-webex

*Automatically create webinars in Webex Webinar based on information in a Sharepoint List*


It is easy to use:

Collaborate with your team on webinar planning. When ready for creation, check **Create=yes**

.. image:: docs/images/sharepoint-prepare.gif
    :width: 1500
    :alt: Sharepoint Lists screenshot showing preparation of webinar data and marking webinars for creation

Schedule all webinars with one bot command

.. image:: docs/images/bot-schedule.gif
    :width: 854
    :alt: Webex bot screenshot showing clicking the Schedule Now button

Webinars are created

.. image:: docs/images/sharepoint-complete.gif
    :width: 1500
    :alt: Sharepoint Lists screenshot showing the scheduled webinars details appearing

If need to change title, description, or reschedule, run the bot command again. You can also run it on a schedule.


Features
--------
This automation ties together three different services: Sharepoint, Webex Meetings/Webinars and Webex Messaging bot. It helps a lot if you are running many webinars, especially in series, especially with multiple people collaborating on them.

This automation supports:

- Create and update Webex Webinars based on information in a Sharepoint list
- Reports status via bot to a Webex space
- Control with Webex bot adaptive cards
- Creation can be triggered by bot command or by schedule
- Customizable webinar parameters
- Attendee link, host key and registrant count updated into the Sharepoint list


How it works
------------

- Collect all webinar information in a Sharepoint list, one webinar per row. Include details like webinar title, description, date and time, hosts, panelists etc. The list can be shared by multiple people for teamwork.
- Check out individual webinars for creation by changing the ``Create`` field to ``yes/True``. Save the changes.
- Mention the @bot in the Webex room and click ``Schedule now`` button.
- The scheduling will be triggered and the bot will report back after some seconds (or minutes, depending on your amount of webinars).


Get Started
-----------

This automation requires a few things to be set up. Look for details in `Get Started <docs/get_started.rst>`_


Contribute
----------

Feel free to fork and improve.


Support
-------

This automation is offered as-is.
