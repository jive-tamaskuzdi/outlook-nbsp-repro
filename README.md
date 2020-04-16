# About

This is a reproduction repo for issue https://github.com/OfficeDev/office-js/issues/1088

# Steps to reproduce

## Preparations
1) Have OSX Catalina installed
2) Have up-to-date Outlook app installed
![Image of outlook version](outlookversion.png)
4) Send an email that has `hello&nbsp;world` in it, or copy the output from here: https://codepen.io/jive-tamaskuzdi/pen/BaojeJe

# Using the app
4) [Sideload manifest.xml](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing) to outlook app 
5) Install dependencies with `npm ci`
6) Start the app with `npm run dev-server`

# Expected outcome
When I have `hello world` highlighted and I click it, the contextual addin with the taskpane.html appers.

# What actually happens
- I have both `hello world` texts highlighted, which is good
- When I click `hello world` in the first line, the contextual addin with taskpane.html loads fine, this is still good
- When I click `hello world` in the second line, the contextual addin doesnt show and nothing happens.