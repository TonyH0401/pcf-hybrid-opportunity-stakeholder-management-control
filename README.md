# Introduction

A **hybrid** PCF control with FluentUI for managing Opportunity & Stakeholder relationship in **Sales Configuration** Solution - **Sales Hub** Customize Application.

# Getting Started

Install the dependencies: `npm install`.

Build the PCF: `npm run build`.

Run the PCF in local (with watch like `nodemon`): `npm run start:watch`.

Publish/Deploy the PCF control: `pac pcf push --publisher-prefix <your publisher prefix>`.

- https://learn.microsoft.com/en-us/power-apps/developer/component-framework/import-custom-controls#deploying-code-components.
- Make sure you are in the correct account and environment, star the account and environment to select them.
- For managed solution, remember to change the version in the manifest file before you build and deploy.

# Development Note

_Ordered from newest to oldest_

These are the files you will modify and need to be careful about: `manifest`, `manifest type` and `index`. The `manifest` file is the file where you define input for the data input field, add tools with `webAPI`, `utility`, etc.

In `index.ts` there are 4 parts:

- `init()` runs once when the app is initialized to initialize the properties.
- `updateView()` runs when any data is updated
- `getOutputs()` gets the output...tbc
- `destroy()` ...tbc

`container` represent the UI, there is only 1 container, but a container can have multiple childs. If you try to have a container append child multiple div then try to change value, it will only change value of the first div only. To fix this, you need to append multiple div and have those div append a field used for edit information.

`context` gets data from the context or data input section. Access them through `parameters.<name>.raw`
