#!/usr/bin/env node 

/* 
 Did some research and found:
      • Outlook email messages are stored in: "/Users/<<USERNAME>>/Library/Group\ Containers/UBF8T346G9.Office/Outlook/Outlook\ 15\ Profiles/Main\ Profile/Data/Message\ Sources/0"
      • These are easier to read as they are plain text and are more uniform than the OLM file you can export from Outlook.
*/

// Dependancies
const { Console, error } = require('console');
const fs = require('fs');
const os = require('os');
const chalk = require('chalk');


// Arrays
const OmittedAddresses = ['noreply', 'no_reply', 'donotreply', 'no-reply', 'support', 'newsletter', 'accounts', 'notifications', 'newsletter', 'no.reply'] // Array of strings that we want to omit from the list of addresses.
const OmittedChars = ['%', '!', '#', '$', '%', '\\', '\'', '*', '/', '=', '?', '^', '`', '{', '|', '}', '~'];
var Addresses = Array(); // Array to store the email addresses that are detected.

// Start time
var startTimestamp = new Date(); // Current time, log this so that we can report on how long this took.

//Reg-Ex:
const validEmailTest = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/; // Email REGEX
const validUUIDTest = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}$/i; // UUID REGEX
const illegalChars = /[<>#+=\/±§\(\):]/g; // Illegal email chars

const sortArray = false; // Change to TRUE if you want the resulting file to be sorted alphabetically.
const username = os.userInfo().username; // Get the username of the user who is logged in, this way it will target their outlook folder.
var parentPath = '/Users/'+ username +'/Library/Group\ Containers/UBF8T346G9.Office/Outlook/Outlook\ 15\ Profiles/Main\ Profile/Data/Message\ Sources/' // Store the path of the folder that holds all of the email messages.

try{
    var folders = fs.readdirSync(parentPath); // Find out all of the folder names that reside in the parentPath
}
catch{
    throw new Error("Error locating message sources folder for Outlook"); // If the directory doesn't exist then throw an error to stop execution
}

folders.shift(); // Remove the first folder name because it's a fodler we're not interested in.
folders = DeDupeArray(folders); //Sort the array.

var totalFolderCount = folders.length; // The total number of folders that need to be searched through.
var folderCount = 0; // The current folder being worked on.
var totalFileCount = 0; // Total number of files that the application has worked through so far.
var LastExportedLength; // The length of the array at the time we last exported it.

try{
    folders.forEach(folderName => {
        var folderPath = parentPath + folderName + '/';
        folderCount++;
        var folderPercentage = (folderCount/totalFolderCount)*100;

        
        var files = fs.readdirSync(folderPath);
        totalFileCount += files.length;
        var fileCount = 0;

        var currentTime = new Date();
        var elapsedTime = new Date (currentTime.getTime() - startTimestamp.getTime() -  1*60*60*1000); // Minus 1 hour because it seems to add one on
        elapsedTimeFormatted = elapsedTime.toLocaleTimeString();
        
        console.clear();
        console.log(OmittedAddresses);
        console.groupCollapsed(chalk.bold.green('Progress    ' + Math.round(folderPercentage) +'%' + '    |     Address count      ' + Addresses.length + '   |   Time elapsed       ' + elapsedTimeFormatted + '    |   Emails scanned     ' + totalFileCount ));
        console.group('Folder   ' + folderName + ' (' + files.length + ' files)');

        files.forEach(fileName => {
            var skipEmail = false; 
            var AddressBefore = Addresses.length;
            console.group('File    ' + fileName);
        
            var filePath = folderPath + fileName;
            fileCount++;
            fileData = fs.readFileSync(filePath);
            

            var stringData = fileData.toString();
            var lineCount = 0;

            /*try{ // Creating recursive function to trim down the file so that we only loop from the start of the file to the last line that has a valid email address.
                var splitData = TrimFile(stringData);
            }catch(e){
                throw new Error(e);
            }*/

            if(splitData == undefined){
                var lastAtPos = stringData.lastIndexOf('@');
                var lastReturnPos = stringData.indexOf('\r', lastAtPos);
                var splitData = stringData.substring(0, lastReturnPos).split("\r"); // Split the data at the \r char but only for the piece of the string that contains @ symbols.
            }

            if(splitData != undefined){
                splitData.forEach(dataItem => {
                    if(dataItem.includes('> Begin forwarded message:') || dataItem.includes('-----Original Message-----') || dataItem.substring(0, 1)  == '> ' || dataItem.includes('X-Spam-Status: Yes') || skipEmail){ 
                        skipEmail = true; // Email is forwarded and we shouldn't collect the email address.
                        return;
                    }
                    else if(!skipEmail){
                        //IF the current line includes: "To", "CC", "BCC", "From", or "<" and ">" then check if we've got an email on that line.
                        if (dataItem.includes('To:') || dataItem.includes('CC:') || dataItem.includes('BCC:') || dataItem.includes('From:') || (dataItem.includes('<') && dataItem.includes('>'))) {
                            if (dataItem.includes('@')) {
                                emailAddress = extractEmail(dataItem);
                                if(emailAddress != undefined){
                                    try{
                                        var lengBefore = Addresses.length;
                                        Addresses.push(emailAddress); // Add the email address to the Addresses[] array
                                        if(Addresses.length > lengBefore) Addresses = DeDupeArray(Addresses); // Deduplicate and sort the Addresses[] array
                                        if((Addresses.length % 250) === 0 && LastExportedLength != Addresses.length){ // If multiple of 250
                                            var result = ExportArray(Addresses); // Export the email addresses gathered to the text file.
                                            if(result){
                                                console.log(chalk.green('Exported ' + Addresses.length + ' to \'EmailAddresses.txt\' on your desktop.'));
                                                LastExportedLength = Addresses.length;
                                            }
                                        }
                                    }catch(e){
                                        console.error(e); // Error
                                    }
                                }                            
                            }
                        }
                    }
                    lineCount++; // Increment the counter
                });
            }else{
                console.log(stringData);
            }
            var addressesAddedThisFile = Addresses.length - AddressBefore;
            if(addressesAddedThisFile > 0) console.log(chalk.green('Found ' + (Addresses.length - AddressBefore) + ' new email addresses'));
            // console.log( chalk.visible( lineCount + ' lines'));
            console.groupEnd();
        });
        console.groupEnd();
        console.groupEnd();

    });
} catch(e){
    console.error('There was an error reading the files.(' + e + ')' ); // Catch it incase it stops for any reason.
}

ExportArray(Addresses); // Export the content of the Addresses[] array to a text file.

function extractEmail(string){
    // This was my original attempt at making this detect email.
   try{
        string = string.toString().toLowerCase();                           // Convert the parameter to lowercase and string.
        var emailStartPos = string.indexOf('<') + 1                         // Find the first position of the '<' character
        var emailEndPos =   string.indexOf('>', emailStartPos);             // Find the first position of the '>' character starting from 'emailStartPos'
        var emailAddress =  string.substring(emailStartPos, emailEndPos);   // Get a substring, starting from 'emailStartPos' and ending at 'emailEndPos'
        var splitEmail = emailAddress.split('@');
        var emailPrefix = splitEmail[0];
        if (emailPrefix === undefined) return;                               // Return if there's no prefix as the email address cannot possibly be valid.
        var validUUID = validUUIDTest.test(emailPrefix);                    // Test to see if the first part of the email is a UUID.
        var validEmail = validEmailTest.test(String(emailAddress).toLowerCase()); // Test to see if the email address is a valid one.
        var EmailWithoutNumbers = emailPrefix.replace(/[0-9]+/g,'');

        if(validEmail && !validUUID){
            var OmitAddress = OmitAddressYN(emailPrefix); // Run to see if the email address contains any values we want to omit.
        }else{
            var OmitAddress = false; // Default value
        }

        if (validEmail && !validUUID && (EmailWithoutNumbers.length + 3 >= emailPrefix.length) && !OmitAddress) {
            return emailAddress; // If the email is a valid email and is not a UUID then return the address.
        }
        else{
            return; // Error validating email.
        }
    }catch{
        return; // Error extracting email.
    }
    
}

function ExportArray(array){
    try{
        var exportPath = '/Users/'+ username +'/Desktop/' + 'EmailAddresses.txt';  // User's desktop path to export 'EmailAddresses.txt'
        var arrayString = array.join('\n'); // New string that has the entire array seperated by ¶
        fs.writeFileSync(exportPath, arrayString); // Export the contents of 'arrayString'
        return true;
    }
    catch{
        throw new Error('Error exporting emails to text file.'); 
    }
    
}

function OmitAddressYN(emailPrefix){
    var omit = false; // Result of the function

    //AddressPrefixLoop: 
    OmittedAddresses.forEach( prefix => { // Filter through and remove any emails that are 'noreply', 'accounts' etc....
        var containsPrefix = emailPrefix.includes(prefix);
        if(containsPrefix){
            omit = true;
            //break; //AddressPrefixLoop; // Exit loop
        }
    });
    
    //CharacterContainsLoop: 
    OmittedChars.forEach( char =>{ // Filter through and remove any emails unwanted characters
        var containsChar = emailPrefix.includes(char);
        if(containsChar){
            omit = true;
            //break; // Exit loop
        }
    });
    return omit;
}

function DeDupeArray(array){
    try{
        let deDuplicated = [...new Set(array)];
        if(sortArray)
            var sortedArray = deDuplicated.sort();
        return deDuplicated;
    }catch{
        console.error('Error deduplicating array.');
    }
}

function TrimFile(data){ // Function to trim the tail end off of large files.

    var lastAtPos = data.lastIndexOf('@'); // Find the position of the last @ symbol.
    if(lastAtPos === -1 || data.length === 0) return data; // Return, there can't possibly be an email without an @ symbol.
    var lastReturnPos = data.indexOf('\r', lastAtPos); // The position of the next closest \r after the last @ symbol
    if (lastReturnPos < lastAtPos && lastAtPos != -1){ // If the @ is after the ¶ or the @ pos is -1 then we have an error.
        
        if(lastReturnPos != -1) // if we have a return, use the position of that to split the data
            var splitPos = lastReturnPos;
        else
            var splitPos = data.length; // Otherwise just get the whole thing.

        var splitData = data.substring(0, splitPos).split('\r');
        return splitData;
    }else{
        var splitData = data.substring(0, lastReturnPos).split("\r"); // Otherwise split it but only for the part of the string we want.
    }

    var lastItem = splitData[splitData.length - 1]; // Get the last item in the array
    var email = extractEmail(lastItem); // Check if it contains an email address
    if(email === undefined || !validEmailTest.test(email)){ 
        splitData.pop(); // If it doesn't then remove the last item in the array
        data = splitData.join('\r'); // Rejoin the data.
        var result = TrimFile(data); // Run the process again on the now smaller string.
        return result; // Return the result.
    }else{
        return splitData; //Returns the result to 'result'
    }
}