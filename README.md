# leaveapp

This is a skimmed down leave application app, which can mostly be utilised by Startups.

It takes input through a no-code app frontend, puts the entries in google sheets, and then from google sheets it goes to a common calendar(it can either be organisational calendar, or function calendar). Slack integration is done for google calendar events which notifies of new leaves, edited leaves, deleted leaves.

The idea of the app is to allow employees to apply for new leaves, edit already existing leaves or delete future leaves.

I had used no-code solution provider(glideapp) for creating the front-end. It has some bit of validation(but some crucial elements aren't there).

The backend is purely google sheets, google calendar and app script.(App Script is very powerful)

What does the app does?

1. Allows to create new leave application. Send email to users when leave application is successfully added to common calendar, or all required fiels aren't filled, or start leave date is greater than end leave date. New leave when created is updated on common calendar as well.
2. It allows users to edit already existing leave, which is updated on common calendar as well. Already existing leave is only editable if end date is greater than today/now(time)
3. It allows users to delete already existing leave. ALready existing leave can only be removed if start date is greater than today/now(time).

Feel free to make changes on the backend script, and make it better(sending emails to managers about their reportees).

I haven't done any performance on the "for loops" used. If someone can optimise it, will be very happy.

I am bad at writing recursion queries. Will be very hapy if someone can do it.

A glimse of how the front-end looks like.

Takes your gmail account to login to the app, and shows only your leaves.
<img width="383" alt="Screenshot 2022-03-28 at 3 08 51 PM" src="https://user-images.githubusercontent.com/59099012/160371049-a274df8f-df3a-4b90-aed8-95eddbb86049.png">


Ading a New Leave
<img width="392" alt="Screenshot 2022-03-28 at 3 07 39 PM" src="https://user-images.githubusercontent.com/59099012/160370705-ede5158f-5f6a-47b5-8306-9263500179e0.png">

Deleting or editing a leave
<img width="392" alt="Screenshot 2022-03-28 at 3 11 30 PM" src="https://user-images.githubusercontent.com/59099012/160371395-b2a27ecc-543e-4014-a141-61181ebb4c61.png">

Upon doing a CTA event(add new leave, edit already existing leave, delete future leave), google sheets is updated. Any edit on google sheet triggers the function "OnChangeTrigger", which looks for the kind of event triggered, and accordingly calls the respective function.

The respective functions update data from google sheets on google calendar.

You can reach out to me at deepankarboro@hotmail.com for more information.

