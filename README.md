# Contributing to Automation Anywhere Bots

Thank you for taking the time to contribute to Automation Anywhere Bots. 

The following are a list of guidelines and best practices for contributing changes to [Bots](https://docs.automationanywhere.com/bundle/enterprise-v2019/page/enterprise-cloud/topics/aae-client/bot-creator/developer-recommendations/cloud-build-reusable-bots.html) and [Packages](https://docs.automationanywhere.com/bundle/enterprise-v2019/page/enterprise-cloud/topics/aae-client/bot-creator/developer-recommendations/cloud-build-reusable-package.html) which are intended to be available on the Automation Anywhere Bot Store. 

### Table Of Contents

- Code of Conduct
- AAI Free Source Code License
- Bot Development Best Practices
- Package Development Best Practices
- How can I contribute?


### Code of Conduct

Participants in Automation Anywhere free source code projects are expected to adhere to a Code of Conduct.  Please read the full text [here](https://github.com/AutomationAnywhere/AAI-Botstore-Open-Source-Bots/blob/master/AAI%20Contributor%20Covenant%20Code%20of%20Conduct.md) so you understand the boundaries of participation in Automation Anywhere free source code projects.  

### AAI Free Source Code License

This Automation Anywhere Bot is issued under the AAI Free Source Code License. License information and text is included within Bot source code, as well as in this Bot’s Git repo.
The AAI Free Source Code License is based on the MIT License, but with additional terms that: 

(1) require the Bot source code and any works based on the Bot source code to be used only with Automation Anywhere platforms; and
(2) state license termination conditions. 

### Bot Development Best Practices 

#### Clearly Defined Inputs/Outputs/Prerequisites
Bots developed for re-usability should have clearly defined feature(s) with clearly defined inputs and outputs. Variables which are marked as Input or Output values should have clear, meaningful descriptions so that developers know what data is required, and in what format it should be expected. Additionally, a bot’s output data, type, and possible error returns should be clearly communicated so that developers using this bot as a sub-task can appropriately address the potential returns. Should the sub-task require that expected applications are pre-installed or already open on bot run, those prerequisites should also be clearly communicated. 


#### Single Responsibility Principle
Bots developed for re-usability should follow the single responsibility principle which states each sub-task/component should have responsibility over a single part of the functionality of the overall bot and that responsibility should be entirely encapsulated by that sub-task/component. In this way, bots designed and developed to be used as sub-tasks can find maximum re-usability across use cases. Tasks that are the same (or very similar) across multiple processes mean that developers can create a sub-task once and re-use it across multiple use cases - speeding up the development process and preventing re-work.

#### Appropriate Task Cleanup
Any applications, files, or windows which are opened by a bot/sub-task should also be appropriately closed by the bot/sub-task - which includes handling these items on success as well as failure. Likewise, it's important that any applications opened prior to calling a bot/sub-task are not inadvertently closed. Additionally, it's important to make sure that the application that your bot/sub-task depend on are already in place and installed in the expected locations. Consider validating the presence of said applications before attempting to launch them. Finally, there will be instances where a bot/sub-task purposely opens, and leaves open, an application - make sure these scenarios are carefully documented in any documentation so that the calling bot knows to be responsible for closing any application windows opened by a sub-task.

#### Error Handling
Beyond successfully completing the task at hand, a bot’s primary goal is to make sure that any failure or exception is handled gracefully. As such, each task and sub-task should have it's own error handling. An unhandled exception in a sub-task can cause issues for a parent task so it's especially important that bots designed to be used as sub-tasks appropriately address errors using the try, catch, and finally actions in A2019 . Use try/catch/finally blocks at the root level of every bot. Also consider using the same approach around specific operations your bot is taking that may be more prone to error. Consider using try/catch blocks inside of a loop should you want to try an operation multiple times before reporting a failure. 

#### Portability
When designing and testing a bot (or bot designed to be used as a sub-task), it's important to consider that at some point, the bot is likely to be executed from a machine other than the machine it was developed on. Keep in mind which values within a local file path, a network share, or even a window title may need to be variablized so that the bot can appropriately execute when being run from another machine. For things like window titles, this also means using wildcards in the window title name which may have references to the currently logged in user, the environment that is being used for development, or an edition of the target application which may change at some point.

##### Instead of:
Salesforce - Professional Edition - Internet Explorer 
##### Consider Using:
Salesforce - * - Internet Explorer

#### Prompts, Message Boxes, and Possible Infinite Loops
Testing is essential when developing a bot designed to be used as a sub-task. As a part of that testing, make sure that any prompts or message box commands have been removed/disabled from the code in any bot that is intended to run as an unattended automation. A sub-task that interrupts the flow of a calling unattended, parent-bot is not helping the process at all. Additionally, carefully analyze scenarios in which the sub-task being designed could be called - be sure that beyond error handling, loops have a definite end, and that processing wouldn't completely hang up due to a logic error in the designed sub-task.
Note that the removal of such prompts and message boxes applies to bots which are intended to be run in unattended mode. If you’re building bots for attended automation, such message boxes and prompts are often very reasonable or required for bots to run as expected. 

#### Variable Naming
Follow established/communicated practices for variable naming. Automation Anywhere suggests following a Hungarian Notation in A2019 in which each variable is prefixed with a lowercase character(s) indicating the variable’s type. 
Sample Variable Name | Type
------------ | -------------
aMyVariableName | Type: Any 
sMyVariableName | Type: String
nMyVariableName | Type: Number
bMyVariableName | Type: Boolean
lMyVariableName | Type: List
tMyVariableName | Type: Table
rMyVariableName | Type: Record
dMyVariableName | Type: Dictionary
dtMyVariableName | Type: DateTime - uses 2 lowercase letters 
wMyVariableName | Type: Window

#### Credential Vault
The included Credential Vault in the Enterprise Control Room can be used to appropriately store sensitive data such as usernames, passwords, API keys, Tokens, etc. These sensitive values should never be hard-coded into a bot or a bot designed to be used as a sub-task, as their hard-coded storage in a bot introduces a security risk. Instead, create a Locker in the Control Room Credential Vault to securely store these sensitive credentials and fetch them as needed by referencing the credential and attribute when developing the bot. This allows developers to create bots that consume API’s, perform logins, etc. without the need for a bot consumer to directly hard-code the needed credentials within the bot itself. When credential vault values are required to properly use/consume a bot, make sure that all locker names and credentials are clearly called out in the documentation. This could also be extended to include information about how to obtain such credentials (in the case of consuming a 3rd party API which requires registration for an API key/token).

#### Designing for Testability
The Test Driven Design development process starts with writing a test case as each new feature of an application (or bot) is added. These tests should be designed in a way that define the specific function, and the test case should be written in a way that validates that the feature has been added successfully. With the application of the single responsibility principle and designing with re-usability in mind, the result will be many smaller tasks - which can be tested alone in a unit-test fashion. As you’re creating bots for re-usability, design them in a way that they can be tested independently of other sub-tasks - and recognize that as an added benefit, changing one sub-task should have limited impact on other tasks within the process. While this approach may not be possible in all use-cases, this perspective can help make an automation much easier to maintain and deploy.

#### Comments and Steps
Comments allow for developers to put human readable descriptions within their bots so that bot consumers and other developers have an idea of what each section of the code is doing. Commenting for Bot Store submitted bots helps other developers know what each block of code/called sub-task is for, and may clue them in to where they may need to make changes for customizing the bot. Also keep in mind that not all bot developers who may be downloading and using a Bot Store bot have the same development background as the vendor creating the Bot Store submission. Well thought out comments allow all developers to more quickly understand the purpose/function of given code blocks.
Beyond commenting for Bot Store submissions, commenting makes code maintenance easier as the section descriptions clue the developer in on where to look for making changes, and ideally makes for quicker resolution to identified issues. Finally, comments on bots that are a work in progress can be a helpful way of creating placeholders for future work. Consider using a TODO command as a reminder to add logic to a specific area of the bot…making sure such comments are updated/removed once said work is completed.
With A2019 also comes the introduction of the Step action, which allows developers to organize their code into logical groupings that greatly improve readability and organization. Consider creating a rough outline of the major objectives of your bot by using empty, labeled step actions. Once completed, go back through each step, completing the logic for the step as development progresses. In this way, the logic can be cleanly organized into logical groupings and specific functionality can quickly be found/updated as needed. 

#### Logging
When bots and sub-tasks are running unattended on any number of bot runners, identifying problems without logging can be like finding a needle in a haystack…in the dark. Software developers, support teams, and bot owners rely on logs to be able to understand where their automations may be running into issues, and to diagnose said problems. At the very least, bots should be logging errors so that developers can understand when an error occurred and hopefully get some helpful insight into what happened as they work towards resolving the issue. 
Assuming the automation will not be dealing with on-screen personally identifiable information (PII), screen captures, as a part of your logging/error handling process, can be an effective approach to understanding the situation in which the bot/sub-task faced an error.

The [A2019 Bot Store Template](https://github.com/AutomationAnywhere/A2019_Bot_Store_Bot_Template) automatically provides its own log management and logging for errors. In addition to error logs, consider using additional logging measures to help with creating a full audit history of everything a bot/sub-task has done. These logs could be stored in additional log files to include audit, debug, and performance information about the bot. This might include:
•	Main bot start & end time
•	Sub-task start & end time
•	The completion time of specific milestones defined within your bot
•	Number of transactions received in an input file
•	Number of successfully processed/failed transactions


### Package Development Best Practices

#### Know Your Incoming Data
When setting up fields that your Action Package will need from the user, be specific in setting the attribute type to limit the kinds of data that your package will receive. There are 34 attribute types defined in the Java Docs - make sure to review those during your package build so you can select the appropriate field types. Taking in all data as a string means that your code has to do a lot of work possibly converting from one data type to another - so be as specific as possible on limiting the input to reduce the burden of checks that need to be done once the data has come in.
Additionally, consider promoting bot building best practices in the way that your package takes its inputs. For example, if your package will be making API calls on behalf of the bot, make sure that the AttributeType of the action’s input field for API key or Token is set to CREDENTIAL. In this way, users will be encouraged to use a value stored in the Credential Vault for sensitive input data that may be needed by the package. 

#### Use Labels Appropriately
In the CommandPkg annotation, pay close attention to the use and application of the different labels, node_labels, and descriptions that are being used. These labels shouldn’t be full descriptions of your action, just brief 1-3 word labels so a bot developer knows what they are using. Pay close attention to the way that labels/node labels are used on the default Action Packages in an effort to replicate the same naming style. Each action is a child element of a package, and the action label is displayed along with the package’s icon in the Actions Pane. Names that are longer than they need to be may mean bot developers have to constantly expand and contract the width of that pane - which can make for a frustrating user experience. 
Beyond the labeling at the package and action level, it's important to provide guidance to bot developers at the field level within an action. Most fields may be fine with just a simple field label - however you may find that bot developers would benefit from specific guidance on expected input format for certain fields. In those cases, consider using the description parameter for the @Pkg annotation. This will allow for package developers to give additional guidance into the format/requirement/data that should be used for a specific input field.

```@Pkg(label = "API Key", description="Expected format: ACxxxxxxxxxxxxxxxxxxxxxxx")```

```@Pkg(label = "Start Date", description="Date Format as MM/DD/YYYY")```

#### Unit Tests
During package development, it’s important that developers are appropriately creating unit tests to validate that each component and action of their package is working as expected. Unit testing is a method that instantiates a portion of your application, verifying its behavior independently from other components. It’s done to validate the behavior of these individual, testable units – a single class, a single action, etc. – to ensure that the component is working as expected. Unit testing can be helpful in validating that each individual piece of your package is functioning properly, and can also be helpful in detecting defects at early stages of the development process.

#### Error Handling
A responsibility of all bot developers is to include error handling in their bot logic to ensure that, in the case an error occurs, that error is handled gracefully, and that the bot doesn’t experience an error that would prevent the machine from taking other tasks. In that same vein, it's important that the actions which package developers are creating respond back with meaningful errors that can lead bot developers in the right direction for error resolution. As a package developer, two specific things should be kept in mind. The first being that just like the direction given to the bot developer, the action itself should fail gracefully - having appropriate error handling in place (try/catch block for example) to accommodate for an error. Second, it's important to make sure that the error messaging returned to the user includes actionable guidance on resolving the issue which was faced. In package development, that might mean using a multi-catch block to catch specific errors, and the use of BotCommandException to return customized error messages. 
```
//create array of 3 items
int[] myIntArray = new int[]{1, 0, 7};
try {
    //print 4th item in array
    System.out.println(myIntArray[3]);
    //Perform operation on first and second items in array
    int result = myIntArray[0] / myIntArray[1];
} catch (ArrayIndexOutOfBoundsException e) {
    //Throw custom message for IndexOutofBounds
    throw new BotCommandException("The array does have the number of expected items.");
} catch (ArithmeticException e) {
    //Throw custom message on Atithmetic Exception
    throw new BotCommandException("Math Operation Error with " + Integer.toString(myIntArray[0]) + " and " + Integer.toString(myIntArray[1]));
}
```

#### Single Responsibility Principle
A package is a collection of actions. Each action within a package should have a single responsibility and that responsibility should be encapsulated (as best as possible) by that action. That said, don’t try to cram too much logic into a single action just for the sake of having fewer actions. Take an example of creating a package to interface with the A2019 Enterprise Control Room. As opposed to fetching an auth token with every action, the package could have one dedicated action for authenticating and retrieving an authentication token, one action to fetch audit data, and a separate action to deploy a bot. As a package developer following the single responsibility principle, your software is easier to implement, simpler to create unit tests for, and is mostly compartmentalized to avoid unexpected side effects of future changes. From a bot developer perspective, the actions that you offer allow them to customize the way that they use your package within their bots, and can help their bots be as efficient as possible. 

#### Demonstrable Examples
Packages developed for submission to Bot Store, or developed for internal purposes, should additionally have a bot that demonstrates the appropriate use of the package. The obvious benefit of packages is that developers have the ability to expand on the existing capabilities of A2019 and allow package consumers to extend their bot’s capabilities. That said, it’s important to remember that not everyone who may be downloading your package from Bot Store will have the development expertise and product familiarity that you may have. Sample bots (along with proper documentation) are essential to empowering consumers of your package with the knowledge and examples they need to understand its proper use. 


### How can I contribute?

#### Fix any issues and open a Pull Request
Open a new GitHub pull request with the patch. Ensure the PR description clearly describes the problem and solution. Include the relevant issue number if applicable.



#### Reporting Bugs
Before creating bug reports, please check if the issue has already been reported as you might find out that you don't need to create one. When you are creating a bug report, please include as many details as possible. 

##### Note: If you find a Closed issue that seems like it is the same thing that you're experiencing, open a new issue and include a link to the original issue in the body of your new one.

#### Before Submitting A Bug Report
•	Check the [APeople forums](https://apeople.automationanywhere.com/s/topic/0TO6F000000oT3rWAE/bot-store) for a list of common questions and problems.
•	Perform a search to see if the problem has already been reported. If it has and the issue is still open, add a comment to the existing issue instead of opening a new one.

#### Include details about your configuration and environment
•	Which version of Automation Anywhere are you using? 

