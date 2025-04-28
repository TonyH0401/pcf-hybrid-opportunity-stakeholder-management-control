# Introduction 
TODO: Give a short introduction of your project. Let this section explain the objectives or the motivation behind this project. 

# Getting Started
TODO: Guide users through getting your code up and running on their own system. In this section you can talk about:
1.	Installation process
2.	Software dependencies
3.	Latest releases
4.	API references

# Build and Test
TODO: Describe and show how to build your code and run the tests. 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)

# Note

These are the files you will modify and need to be careful about: `manifest`, `manifest type` and `index`. The `manifest` file is the file where you define input for the data input field, add tools with `webAPI`, `utility`, etc.

In `index.ts` there are 4 parts:
- `init()` runs once when the app is initialized to initialize the properties.
- `updateView()` runs when any data is updated
- `getOutputs()` gets the output...tbc
- `destroy()` ...tbc

`container` represent the UI, there is only 1 container, but a container can have multiple childs. If you try to have a container append child multiple div then try to change value, it will only change value of the first div only. To fix this, you need to append multiple div and have those div append a field used for edit information.

`context` gets data from the context or data input section. Access them through `parameters.<name>.raw`