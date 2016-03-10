---
title: |
    Tips and Tricks for Improving Performance and Experience of Office
    Add-ins
...

 

Top tips for experience:
========================

-   Ensure the **user experience** from ‘the very first launch’ to ‘the
    add-in is fully configured’ is easy and intuitive. This is a new
    quality standard that we are enforcing, and is mandatory for all new
    add-ins. 

    -   Be clear at first launch (even before sign-in) what value the
        user gets from the add-in – this can take the form of
        text/imagery/video within the add-in.  This will increase the
        user’s motivation to persist through a sign-in/sign-up process
        so they reap the value of your add-in.

    -   If your add-in requires the user to have an account on an
        external service, provide a link to your website for the user to
        sign-up.  Signing up within the add-in offers the best
        experience with the least drop-off, though we understand that
        (depending on your service) development costs to deliver this
        well may be non-trivial.

    -   Populate your add-in with sample data instead of being
        fully blank.

    -   Consider step by step guided set up.  Display tooltips to guide
        users through set-up.  Don’t assume users have read the manual.

    -   Consider (but don’t fully rely on) supplementary “Getting
        Started” content, such a demo video of your add-in in action,
        links to sample documents or links to your website.

    -   Run usability tests on the user’s first experience in the
        add-in. Add in-app telemetry to see where users drop off.

    -   Tablet/phone app stores have trained users to expect instantly
        usable experiences.  The Office Store will reject add-ins which
        are too hard to set up.

 

-   For Word, Excel, PowerPoint and Outlook, use the most
    recent **manifest schema version 1.1** to get the most out of
    Office.  Failure to use version 1.1. will result in demotion for
    existing add-ins, and rejection for new ones. This policy ensures
    add-in work in Office Online and Office for Mac 2016, and have best
    load-time latency.  

    -   Test your add-in in Word/Excel/PowerPoint Online and Outlook Web
        Access across all supported browsers (Firefox, Edge, IE and
        Chrome).  Test also on Safari on a Mac.

    -   If you take a hard dependency on Office 2016, make sure your
        add-in states this in the requirements section of the schema. 
        Ideally though, have your add-in give a graceful fall-back
        experience for Office 2013 users.

    -   Office for Mac will soon support add-ins. You can download a new
        version of Word for Mac \[Version 15.18 (160109)\] from
        Office365.com or using the Microsoft AutoUpdate app that
        supports the developer preview version of Office add-ins. The
        developer preview version provides an opportunity for you to
        test and validate your add-in before we release Add-ins on the
        Office for Mac platform for consumers. 

    -   Here are a couple of articles to help you get started with
        validating your add-in on the new Word for Mac:

> [Side-load an Office Add-in on
> Mac](https://msdn.microsoft.com/EN-US/library/office/mt154253.aspx)
>
> [Debug Office Add-ins on
> Mac](https://msdn.microsoft.com/EN-US/library/office/mt459579.aspx)

-   If you have an **Outlook** add-in, ensure you add app commands to
    that your add-in has a ribbon presence.  We will imminently release
    new guidance for Outlook add-ins app commands, so please check back
    in a couple of weeks for a blog post with more details before you
    implement this.

-   Make sure your privacy and support links are available at all times.
    Add-ins will be demoted in the store ranking if the URLs
    are unavailable.

 

For best business success, alter your **merchandising**:
========================================================

-   Use the **title and short description** to convince as large an
    audience as possible to click on your add-in.  This prevents
    needless user drop off when user’s cannot find the add-ins that
    would otherwise be valuable. Titles which don’t state the add-in’s
    functionality are at risk of rejection.

    -   Don’t rely on existing users already knowing your brand… Win new
        customers by ensuring your title contains your brand AND a 2-4
        word description of what it is, and that your short description
        contains your value proposition.  Here’s a really simple
        example: Change from ‘Fabrikamation for Office 365’ to
        ‘Fabrikamation – Warehouse management’

    -   Don’t heavily duplicate words across title and description –
        make every word count.

    -   Just as you would do for your website, think about all the
        keywords a user might type

    -   Ask your target audience for what they think of your title and
        short description – it’s the simplest thing to usability test,
        but it’s critically important

 

-   Annotate **screenshots** of your add-in with your value
    proposition.  Screenshots are very heavily used by users to see if
    they want to try the add-in.

    -   We encourage call-outs/imagery in screenshots which point to the
        value users derive from various functions.

 

Take advantage of our new feature – **developer responses for user reviews**:
=============================================================================

-   We understand how important it is for developers to engage directly
    with their customers. With this in mind, we have just released an
    update to our website which enables responses within reviews.

    -   Make sure you're signed in with the account that you used to
        submit the add-in when responding to reviews so that your
        comments get marked with our "App Provider" tag.

    -   Consider reminding users to alter their review if you have
        addressed their issue.

 

Finally, we’d like to share a policy change with you.  We recognise that
as new policies are rolled out, not every developer (with an existing
add-in in the Store) can update their add-in immediately to meet the new
criteria.  We’ve also received questions a few times of why a new
submission must meet all policies when older add-ins do not.  Our new
policy is:

-   **Existing add-ins** in the Store which do not pass newer policies
    can (under most circumstances) stay unchanged in the Office Store
    but will be demoted and will no longer receive prominent placement.

-   All **new and updated submissions** to the Office Store must pass
    all policies.  This will give your add-in a great opportunity for
    prominent store placement.
