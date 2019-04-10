# pptx-Editor

This is Java pptx Editor with GUI in JavaFX.

## Getting Started

### Preparation 

* At first you should have jre 8 or 10 on your computer.

https://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html

* Then add a java path.

https://www.java.com/en/download/help/path.xml

### Edit pptx

This app works only with .pptx Power Point presentations.
You can run this program with **startJava8.vbs** or **startJava10.vbs** scripts which you can find in the main folder of the project.

<img src="https://github.com/IlyaKonka/pptx-Editor/blob/master/doc/mainMode.png" width="400">

To start editing presentations just write the path to presentation folder. (You can also scan the whole disk)
Then write word or words which you would like to replace in the presentations and on the right the new replacement word or words.

Extra functions:

* **Picture mode** - In this mode you can add a text to each picture in each presentation. Style and color configuration of font is also
available. There are two subfunctions. 

*Write many times* - the whole picture will be filled with text.  

*Rotate pictures* - all pictures will be rotated by 180 degrees.

* **Backup** - This mode is needed to create a backup of all presentations which were changed. At first write a path to backup zip file which will be
created. This zip file should have a password. It is necessarily to crate the file with list of the names of all modified presentations. By default this text file will be
created in the same folder as backup zip. You have also a possibility to get zip password and the list with names of presentations on your email.

**To start** the program click Start button. If the edit of presentations was successful, you will see **Done** message in the console. 

### Backup pptx

<img src="https://github.com/IlyaKonka/pptx-Editor/blob/master/doc/backupMode.png" width="400">

To back up the modified presentations use **Back up Mode**.
Write the path to backup zip with his password and the path to file with names of presentations which you had to create in the **Main mode**.
Press **Start** and check the log of console. You can see there which presentations were backed up successfully.


## Additional

It's a maven project. To build it with java 8 uncomment this lines:

```
<source>1.8</source>
<target>1.8</target>
<verbose>true</verbose>
```

And comment this:

```
<source>10</source>
<target>10</target>
<release>10</release>
```

You should also make some adjustments in the **Main.java** file to get correct resize for java 8 or java 10.

