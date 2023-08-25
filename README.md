# Macros for Connexion Client (v. 3.0)
Macros for use in Connexion Client v. 3.0, written in OML (OCLC macro language), an offshoot of BASIC and kind-of cousin of VBA.


If you are writing or editing macros:

'OML for the complete beginner' by Joel Hahn (http://www.hahnlibrary.net/libraries/oml/lessons/index.html) is a good introduction to macros in OCLC Connexion and is freely accessible online. It will also introduce you to general coding basics that will help if you go on to learn another coding language.

You can find a list of commands on OCLC’s website (https://help.oclc.org/Metadata_Services/Connexion/Connexion_client/Connexion_client_basics/Use_macros/List_of_Connexion_client_macro_commands/Connexion_client_macro_commands_A-Z). It doesn’t list not every command that exists in OML but it is a good start, and helps you see what kinds of things macros can do for you.

OML is a version of Softbridge Basic Language (SBL) and is similar to VBA. As such, you can look at VBA solutions online for ideas. Microsoft’s information on VBA (https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference) may be of help to you. Since this is an offshoot of BASIC you can also try to search for commands in that.

Note: Back up your macros if you have edited them. Sometimes after an edit Connexion will cut the bottom of the code off or duplicate a chunk at the bottom. I recommend saving them as text files.

Helpful notes:

* Message boxes have character limits
* If you are using the Dir() function, you must supply two arguments. The second one is technically optional, but it seems to need both in OML. 
