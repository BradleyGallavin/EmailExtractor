#!/usr/bin/env node 
//The shebang!

// Attempt 3 at the email fetching thing

// Did some research and found:
//      • Outlook email messages are stored in: "/Users/bradleygallavin/Library/Group\ Containers/UBF8T346G9.Office/Outlook/Outlook\ 15\ Profiles/Main\ Profile/Data/Message\ Sources/0"
//      • These are easier to read as they are plain text and are more uniform than the OLM file you can export from Outlook.


// Set some initial variables that we're going to need.

const { Console, error } = require('console');
const fs = require('fs');
const os = require('os');

// Arrays
const OmittedAddresses = ['noreply', 'no_reply', 'donotreply', 'no-reply', 'support', 'newsletter', 'accounts', 'notifications'] // Array of strings that we want to omit from the list of addresses.
var Addresses = Array(); // Array to store the email addresses that are detected.

// Start time
var startTimestamp = new Date(); // Current time, log this so that we can report on how long this took.

//Reg-Ex:
const validEmailTest = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/; // Email REGEX
const validUUIDTest = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-5][0-9a-f]{3}-[089ab][0-9a-f]{3}-[0-9a-f]{12}$/i; // UUID REGEX

const username = os.userInfo().username; // Get the username of the user who is logged in, this way it will target their outlook folder.

// Store the path of the folder that holds all of the email messages.
var parentPath = '/Users/'+ username +'/Library/Group\ Containers/UBF8T346G9.Office/Outlook/Outlook\ 15\ Profiles/Main\ Profile/Data/Message\ Sources/'

try{
    var folders = fs.readdirSync(parentPath); // Find out all of the folder names that reside in the parentPath
}
catch{
    throw new Error("Error locating message sources folder for Outlook"); // If the directory doesn't exist then throw an error to stop execution
}

folders.shift(); // Remove the first folder name because it's a fodler we're not interested in.

var totalFolderCount = folders.length;  // The total number of folders that need to be searched through.
var folderCount = 0;                    // The current folder being worked on.
var totalLineCount = 0;

try{
    folders.forEach(folderName => {

        var folderPath = parentPath + folderName + '/';
        folderCount++;
        var folderPercentage = (folderCount/totalFolderCount)*100;
        
        var files = fs.readdirSync(folderPath);
        var totalFileCount = files.length;
        var fileCount = 0;

        files.forEach(fileName => {
            
            var filePath = folderPath + fileName;
            fileCount++;
            fileData = fs.readFileSync(filePath);
            

            var stringData = fileData.toString();
            var lineCount = 0;
            var lastAtPos = stringData.lastIndexOf('@');
            var lastReturnPos = stringData.indexOf('\r', lastAtPos);

            var splitData = stringData.substring(0, lastReturnPos).split("\r"); // Split the data at the \r char but only for the piece of the string that contains @ symbols.
            
            // Creating recursive function to trim down the file so that we only loop from the start of the file to the last line that has a valid email address.
            // var splitData = TrimFile(stringData);
            
            if(splitData != undefined){
                splitData.forEach(dataItem => {
                    //IF the current line includes: "To", "CC", "BCC", "From", or "<" and ">" then check if we've got an email on that line.
                    if (dataItem.includes('To:') || dataItem.includes('CC:') || dataItem.includes('BCC:') || dataItem.includes('From:') || (dataItem.includes('<') && dataItem.includes('>'))) {
                        if (dataItem.includes('@')) {
                            emailAddress = extractEmail(dataItem);
                            if(emailAddress != ''){
                                try{
                                    Addresses.push(emailAddress); // Add the email address to the Addresses[] array
                                    Addresses = DeDupeArray(Addresses); // Deduplicate and sort the Addresses[] array
                                    var currentTime = new Date();
                                    var elapsedTime = new Date (currentTime.getTime() - startTimestamp.getTime() -  1*60*60*1000); // Minus 1 hour because it seems to add one on
                                    elapsedTimeFormatted = elapsedTime.toLocaleTimeString();
    
                                    console.clear(); //Clear and update the console with the current progress.
                                    console.log('Time elapsed       ' + elapsedTimeFormatted);
                                    console.log('Address count      ' + Addresses.length);
                                    console.log('Folders            ' + folderCount + ' of  ' + totalFolderCount +' (' + Math.round(folderPercentage) +'%)');
                                    console.log('Current folder     ' + fileCount + '   of  ' + totalFileCount + ' files in folder ' + folderCount);
                                    console.log('Read lines         ' + totalLineCount);
                                    console.log('Current file       ' + lineCount + '   of  ' + splitData.length + ' lines');
                                    if((Addresses.length % 250) === 0){ // If multiple of 250
                                        ExportArray(Addresses); // Export the email addresses gathered to the text file.
                                    }
                                    
                                }catch(e){
                                    console.error(e); // Error
                                }
                            }
                        }
                    }
                    
                    // Increment the counter
                    lineCount++;
                    if(lineCount === splitData.length){
                        totalLineCount += splitData.length;
                    }
    
                });
            }else{
                console.log(stringData);
            }
        });
    });
} catch(e){
    console.error('There was an error reading the files.\n' + e ); // Catch it incase it stops for any reason.
}

ExportArray(Addresses); // Export the content of the Addresses[] array to a text file.

function extractEmail(string){
    try{
        string = string.toString().toLowerCase();                           // Convert the parameter to lowercase and string.
        var emailStartPos = string.indexOf('<') + 1                         // Find the first position of the '<' character
        var emailEndPos =   string.indexOf('>', emailStartPos);             // Find the first position of the '>' character starting from 'emailStartPos'
        var emailAddress =  string.substring(emailStartPos, emailEndPos);   // Get a substring, starting from 'emailStartPos' and ending at 'emailEndPos'
        var splitEmail = emailAddress.split('@');
        var emailPrefix = splitEmail[0];
        var validUUID = validUUIDTest.test(emailPrefix);                    // Test to see if the first part of the email is a UUID.
        var validEmail = validEmailTest.test(String(emailAddress).toLowerCase()); // Test to see if the email address is a valid one.
        var EmailWithoutNumbers = emailPrefix.replace(/[0-9]+/g,'');
        var OmitAddress = OmitAddressYN(emailPrefix);

        if (validEmail && !validUUID && (EmailWithoutNumbers.length + 3 >= emailPrefix.length) && !OmitAddress) {
            return emailAddress; // If the email is a valid email and is not a UUID then return the address.
        }
        else{
            return;     // Error validating email.
        }
    }catch{
        return;         // Error extracting email.
    }
}

function ExportArray(array){
    try{
        var exportPath = '/Users/'+ username +'/Desktop/' + 'EmailAddresses.txt';  // User's desktop path to export 'EmailAddresses.txt'
        var arrayString = array.join('\n'); // New string that has the entire array seperated by ¶
        fs.writeFileSync( exportPath, arrayString); // Export the contents of 'arrayString'
        console.log('Exported ' + array.length + ' addresses to EmailAddresses.txt');
    }
    catch{
        throw new Error('Error exporting emails to text file.'); 
    }
    
}

function OmitAddressYN(emailPrefix){
    OmittedAddresses.forEach( prefix => { 
        var containsPrefix = emailPrefix.includes(prefix);
        if(containsPrefix){
            return true;
        }
    });
    return false;
}

function DeDupeArray(array){
    try{
        let deDuplicated = [...new Set(array)];
        return deDuplicated.sort();
    }catch{
        console.error('Error deduplicating array.');
    }
}

function TrimFile(data){ // Function to trim the tail end off of large files.
    try{
        var lastAtPos = data.lastIndexOf('@');              // Find the position of the last @ symbol.
        var lastReturnPos = data.indexOf('\r', lastAtPos);  // The position of the next closest \r after the last @ symbol
        if (lastReturnPos < lastAtPos){ 
            var splitData = data.split("\r"); // If there was no \r after that then we need can just split it 
        }else{
            var splitData = data.substring(0, lastReturnPos).split("\r"); // Otherwise split it but only for the part of the string we want.
        }
        var lastItem = splitData[splitData.length - 1]; // Get the last item in the array
        var email = extractEmail(lastItem); // Check if it contains an email address
        if(email === undefined && !validEmailTest.test(email)){ 
            splitData.pop(); // If it doesn't then remove the last item in the array
            data = splitData.join('\r'); // Rejoin the data.
            var result = TrimFile(data); // Run the process again on the now smaller string.
            return result; // Return the result.
        }else{
            return splitData; //Returns the result to 'result'
        }
    }catch(e){
        console.error('Error trimming file.\n' + e)
        return data;
    }
}