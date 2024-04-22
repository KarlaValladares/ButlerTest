/**

 * import necessary libraries 

 * if not installed enter "npm install [library name]" in the terminal

 * troubleshooting:

 * @https://www.npmjs.com/package/puppeteer-chromium-resolver

 * https://stackoverflow.com/questions/74362083/cloud-functions-puppeteer-cannot-open-browser

 * https://github.com/puppeteer/puppeteer/issues/1597

 *

 * resources:

 * command to kill port: kill -9 $(lsof -t -i:8080)

 * https://www.kindacode.com/article/node-js-how-to-use-import-and-require-in-the-same-file/

 * https://brunoscheufler.com/blog/2021-05-31-locking-and-synchronization-for-nodejs

 * https://www.youtube.com/watch?v=PFJNJQCU_lo

 * https://dev.to/pedrohase/create-google-calender-events-using-the-google-api-and-service-accounts-in-nodejs-22m8

 */



// IMPORTS

import { createRequire } from "module";

const require = createRequire(import.meta.url);

import { Mutex } from 'async-mutex';

import { google } from 'googleapis';
import { promises as fs } from 'fs';

import { JWT } from 'google-auth-library';

import pAll from 'p-all';

import {
    publishMessage,
    replyMessage,
    findMessage,
    replyMessagewithrecord,
    getMessageByTsAndChannel,
    uploadImage
} from "./slack.js";


const process = require('process');

const prompt = require('prompt');

const puppeteerExtra = require('puppeteer-extra');

//const fs = require('fs').promises;

const path = require('path');

const { fileURLToPath } = require('url');

const _ = require('lodash');

const PCR = require("puppeteer-chromium-resolver");

const stealthPlugin = require('puppeteer-extra-plugin-stealth');

const schedule = require('node-schedule');

// const credentials = require('./gsa_credentials.json');

//const database_credentials = require('./database_credentials.json');

const readline = require("readline-promise").default;

const SCOPES = ["https://www.googleapis.com/auth/calendar.readonly", 'https://mail.google.com/',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify', 'https://mail.google.com/',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/spreadsheets'];

const dotenv = require('dotenv');

dotenv.config();

puppeteerExtra.use(stealthPlugin());

// CONSTANTS

const mutex = new Mutex();

const { DateTime } = require('luxon');

//Operation Calendar and Operation Email

const TOKEN_PATH = "./token.json";

const CLIENT_SECRET_PATH = "./credentials.json";

const auth = await authorize();

const calendar = google.calendar({ version: "v3", auth });

// const gmail = google.gmail({ version: 'v1', auth })



// People Database Constants

const spreadsheetId = process.env.SHEET;

const ALLTABNAME = "All";

const ID_COLUMN = 1;

const NAME_COLUMN = 2;

const NICKNAME_COLUMN = 3;

const ROLE_COLUMN = 4;

const EMAIL_COLUMN = 5;

const PHONE_COLUMN = 6;

const FAMILYID_COLUMN = 7;

const GRADE_COLUMN = 8;




const STUDENTTIME = 5 * 60 * 1000; //5 minutes

const TUTORTIME = 3 * 60 * 1000; //3minutes

const BOTTIME = 1000; // 1 minute the start time



const MONITORING_EMAIL = process.env.MONITORING_EMAIL;

const OPERATION_EMAIL = process.env.GMAIL; // this should be the operation account

const firefliesAccount = process.env.FF_ACCOUNT // this is for the firefly


const DATE_REGEX = /(\d?\d:\d\d):\d\d\s(\w\w)/;

//const CANCELLED_REGEX = /^\((C|c)ancel/;

const CANCELLED_REGEX = /cancel/gi;

const EMAIL = /^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/;

const ABSENT = 0;

const PRESENT = 1;

const MONITORING_CHANNEL_ID = process.env.SLACK_CHANNEL_ID;

const BOT_ALERT = process.env.BOT_ALERT;


class Event {

    title;

    startTime;

    endTime;

    meetLink;

    guests;

    id;



    constructor(title, startTime, endTime, meetLink, guests, id) {

        this.title = title;

        this.startTime = startTime;

        this.endTime = endTime;

        this.meetLink = meetLink;

        this.guests = guests;

        this.id = id;

    }

}


class Guest {

    name;

    nickname;

    role;

    email;

    phoneNumber;



    constructor(name = "", nickname = "", role = "", email = "", phoneNumber = "") {

        this.name = name;

        this.nickname = nickname;

        this.role = role;

        this.email = email;

        this.phoneNumber = phoneNumber;

    }

}


class AttendanceBot {
    context;
    page;

    constructor(context, page) {
        this.context = context;
        this.page = page;
    }

    /**

     * Initializes instance of Chromium using puppeteer

     * See puppeteer documentation @https://pptr.dev/ 

     * @returns AttendanceBot object

     */

    //https://dev.to/somedood/the-proper-way-to-write-async-constructors-in-javascript-1o8c

    static async initialize(browser) {
        const context = await browser.createIncognitoBrowserContext();
        const page = await context.newPage();

        // Set options for headless mode
        await page.setViewport({ width: 1366, height: 768 });
        await page.evaluateOnNewDocument(() => {
            Object.defineProperty(navigator, 'webdriver', {
                get: () => false,
            });
        });

        return new AttendanceBot(context, page);
    }




    /**

     * Uses login credential to sign into Google account.

     * Goes to session URL and joins.

     * @param {*} meetLink 

     */

    async joinSession(title, meetLink, thread) {

        const page = this.page;

        const filePath = './cookies.json'



        // wait maximum amount of time for page to load

        await page.setDefaultTimeout(0);

        await page.setDefaultNavigationTimeout(0);



        // log into gmail account using .env credentials

        try {

            console.log("logging into gmail...");

            await page.goto('https://accounts.google.com/signin/v2/identifier', { waitUntil: 'load', timeout: 0 });

            console.log("typing email address");

            await page.type('[type="email"]', process.env.GMAIL);

            await page.click('#identifierNext');

            await sleep(5000);



            console.log("typing login password");

            await page.type('[type="password"]', process.env.PASSWORD);

            await page.click('#passwordNext');

            await sleep(5000);

        } catch (e) {

            console.log('logging into gmail failed', e);

            throw Error("unable to log in :(");

        }



        // }



        // save cookies (comment-out when deploying as cloud function because cookies.json becomes a read-only file)

        console.log("saving cookies...");

        mutex.runExclusive(() => saveCookies(page, filePath));



        // go to meet url

        console.log(`going to link @ ${meetLink}...`);

        await page.goto(meetLink, { waitUntil: 'load', timeout: 0 });



        // continue without microphone and camera

        // console.log("dismissing popup");

        // await sleep(2000);

        // const popup = await page.waitForSelector('button[jsname="IbE0S"]', { visible: true });

        // await popup.click();

        // await popup.dispose();



        // join meet

        console.log(`joining ${title}...`)

        await sleep(2000);

        let count = 0;

        while (await this.isInMeet() === false && count < 5) {

            try {

                const joinButton = await page.waitForSelector('button[data-idom-class="nCP5yc AjY5Oe DuMIQc LQeN7 jEvJdc QJgqC"][jsname="Qx7uuf"]', { visible: true, timeout: 5 * 1000 });

                await joinButton.click();

                await joinButton.dispose();

            } catch (e) {

                console.log("retrying join...")

            }

            count++;

        }

        if (await this.isInMeet() === false) {

            // {"Kisun": U03MF2SAXMW, "Kwiseon": U03S548N206, "Grace": U04H80KMGRW, "Hyewon": U054EFX4FAM}

            sendMessage(thread, "<@U05M0P95AGZ> Unable to join 😔. Please check script.");

            throw Error(`unable to join - ${title}`);

        }



        // open people tab

        console.log("opening people tab...");

        await page.evaluate(() => {

            document.querySelectorAll('[class="VfPpkd-Bz112c-LgbsSe yHy1rc eT1oJ JsuyRc boDUxc"]')[1].click();

        });

        console.log(`${title} joined!`);

        await sleep(1000);

    }



    async autoAdmit(interval) {

        const page = this.page;



        const elements = await page.$$('[class="VfPpkd-vQzf8d"]');

        elements.map(async element => {

            const text = await element.getProperty("textContent");

            const json = await text.jsonValue();

            if (json.toString() === "Admit") {

                await element.click();

            }

        });

        await sleep(interval)

    }


    // Post contact information for guest.
    async postContactGuest(thread, guestList) {
        let slackThread = thread;
        let logRecordsSet = new Set(); // Use a Set to store unique entries

        const PEOPLEDATABASE = await getGoogleSheetsData(spreadsheetId, ALLTABNAME);

        for (let guest of guestList) {
            // Skip checking monitoring email for the moment is my account
            // we will need to replace it by 'ModernSmart Team'
            if (guest.email == OPERATION_EMAIL) { continue; }

            let peopleData = getPersonFromData(PEOPLEDATABASE, guest.email, EMAIL_COLUMN);

            for (let personData of peopleData) {
                if (personData) {
                    var invitedName = personData[NAME_COLUMN - 1] || "No name";
                    var phoneNum = personData[PHONE_COLUMN - 1] || "";
                    var role = personData[ROLE_COLUMN - 1] || "No role";

                    // Construct the record
                    let record;
                    if (role == "parent" && phoneNum == "") {
                        record = `${capitalizeFirstLetter(role)}: ${guest.email}`;
                    } else if (role == "parent" && phoneNum != "") {
                        record = `${capitalizeFirstLetter(role)}: ${guest.email}, ${phoneNum}`;
                    } else if (role != "parent" && phoneNum != "") {
                        record = `${capitalizeFirstLetter(role)}: ${invitedName}, ${guest.email}, ${phoneNum}`;
                    } else {
                        record = `${capitalizeFirstLetter(role)}: ${invitedName}, ${guest.email}`;
                    }

                    logRecordsSet.add(record);
                } else {
                    // If personData is not defined, add the email directly
                    logRecordsSet.add(guest.email);
                }
            }
        }

        // Convert the Set back to an array and sort based on role
        const logRecordsArray = Array.from(logRecordsSet).sort((a, b) => {
            const roleA = a.substring(0, a.indexOf(':'));
            const roleB = b.substring(0, b.indexOf(':'));
            return roleA.localeCompare(roleB);
        });

        // Join the array into a string
        const logRecords = logRecordsArray.join('\n');

        sendMessage(slackThread, logRecords);
    }



    /**

     * When inside session, takes attendance by reading guest list

     * @param {*} meetLink 

     */


    async takeAttendance(thread, interval, guestList, eventInfo, fileName) {


        const page = this.page;

        let log = "";

        let startTime = eventInfo.startTime
        let endTime = eventInfo.endTime

        let JSONwithTS = fileName

        let id_record = fileName.split('_')[1];

        let attendanceScreenshot = `attendace_${id_record}`;

        let attendaceKeyinfo = "absenceTs";

        let absenceKeyTutor = "attendaceTuTorTs";

        let absenceKeyStudent = "attendaceStudentTs";

        let absenceKeyBot = "attendaceBotTs";

        // Scrape display names

        const targets = await page.$$("[class='zWGUib']");

        const promises = await targets.map(async element => {

            const text = await element.getProperty("textContent");

            const json = await text.jsonValue();

            const name = json.toString();

            if (name.match(EMAIL)) { // Absent guests will show up in meeting as their emails (only if guest list is visible)

                return { name, status: ABSENT };

            } else {

                return { name, status: PRESENT };

            }

        });

        const guests = await Promise.all(promises); // List of displayed guests 
        const { uniquePresent, uniqueMissing } = await createListGuest(guests, guestList);

        let presentStudentFlag = await readJSONValue(JSONwithTS, 'missingStudentFlag');
        const studentPresent = uniquePresent.some(user => user.role === 'student');
        if (presentStudentFlag == 0 && studentPresent) {
            await updateTsValue(JSONwithTS, 'missingStudentFlag', presentStudentFlag + 1);
        }

        let timeRightNow = new Date()

        const startTimeMeeting = new Date(startTime);
        const endTimeMeeting = new Date(endTime);
        const meetingDuration = (endTimeMeeting - startTimeMeeting) / (1000 * 60);
        const after3mins = new Date(startTimeMeeting.getTime() + (2 * 60 + 59) * 1000);
        const after5mins = new Date(startTimeMeeting.getTime() + (4 * 60 + 59) * 1000);
        const after6mins = new Date(startTimeMeeting.getTime() + (5 * 60 + 59) * 1000);
        const after10mins = new Date(startTimeMeeting.getTime() + (9 * 60 + 59) * 1000);
        const after15mins = new Date(startTimeMeeting.getTime() + (14 * 60 + 59) * 1000);
        const after16mins = new Date(startTimeMeeting.getTime() + (15 * 60 + 59) * 1000);
        const after30mins = new Date(startTimeMeeting.getTime() + (29 * 60 + 59) * 1000);
        const after44mins = new Date(startTimeMeeting.getTime() + (43 * 60 + 59) * 1000);
        const after45mins = new Date(startTimeMeeting.getTime() + (43 * 60 + 59) * 1000);

        let screenshotRecord = await readJSONValue(JSONwithTS, 'postscrenshot');

        if ((after6mins < timeRightNow && screenshotRecord == 0) ||
            (after16mins < timeRightNow && screenshotRecord == 1)) {
            let imagePathNameAttendace = `./record/${attendanceScreenshot}.png`;
            await page.screenshot({ path: imagePathNameAttendace, fullPage: true });
            await uploadImage(MONITORING_CHANNEL_ID, imagePathNameAttendace, thread);
            await sleep(30000)
            await updateTsValue(JSONwithTS, 'postscrenshot', screenshotRecord + 1);
            await fs.unlink(imagePathNameAttendace);
            console.log('Picture deleted')
        }


        log += await createLog(uniquePresent, uniqueMissing);

        let replyID = await readJSONValue(JSONwithTS, attendaceKeyinfo);

        let currentMEssage;

        if (replyID != "0") {

            currentMEssage = await getMessageByTsAndChannel(MONITORING_CHANNEL_ID, replyID);
        } else {
            currentMEssage = 'PM'
        }

        let textRecord1 = await removeAMorPM(currentMEssage);

        let textRecord2 = await removeAMorPM(log);

        if (textRecord1 !== textRecord2) {

            sendMessagewithrecord(thread, log, attendaceKeyinfo, JSONwithTS);
        } else {
            // The string representations are the same, so no need to send a message.
            console.log("There is no change in the attendace.");
        }

        let newUniqueMissing = uniqueMissing.slice();
        let mergedObjects = {};

        // Loop through the copied array and merge objects by role
        newUniqueMissing.forEach(obj => {
            const { role } = obj;
            if (role !== 'parent') {
                // Process student and tutor roles
                if (!mergedObjects[role]) {
                    // If the role does not exist in mergedObjects, create an entry
                    mergedObjects[role] = { role, groups: [] };
                }
                // Add the object to the corresponding role in mergedObjects
                mergedObjects[role].groups.push({
                    name: obj.name,
                    nickname: obj.nickname,
                    email: obj.email
                });
            }
        });

        // Handle parent role separately
        const parentObj = newUniqueMissing.find(obj => obj.role === 'parent');
        if (parentObj) {
            if (!mergedObjects['student']) {
                // If 'student' role does not exist, create 'student'
                mergedObjects['student'] = { role: 'student', groups: [] };
            }
            // Add the parent information to the 'student' group
            mergedObjects['student'].groups.forEach(group => {
                group.parent_email = parentObj.email;
            });
        }

        // Handle parent role separately only if 'student' role exists
        if (mergedObjects['student']) {
            const parentObj = newUniqueMissing.find(obj => obj.role === 'parent');
            if (parentObj) {
                // Add the parent information to the 'student' group
                mergedObjects['student'].groups.forEach(group => {
                    group.parent_email = parentObj.email;
                });
            }
        }

        // Remove 'parent' role from mergedObjects
        delete mergedObjects['parent'];


        // Assign merged objects back to newUniqueMissing
        newUniqueMissing = Object.values(mergedObjects);

        if (timeRightNow < after45mins) {
            for (const absentee of newUniqueMissing) {
                if (timeRightNow < after30mins) {
                    await notifyAbsence(
                        absentee,
                        eventInfo,
                        thread,
                        JSONwithTS,
                        absenceKeyTutor,
                        absenceKeyStudent,
                        absenceKeyBot
                    );
                }

                let presentUserEmail;
                let absenteeName;
                let absenteesInfo;
                let studentEmail = await readJSONValue(JSONwithTS, 'sendEmailStudent');
                let studentFlag = await readJSONValue(JSONwithTS, 'missingStudentFlag');
                let reviewIgGroupClass = await readJSONValue(JSONwithTS, 'isGroupClass');
                let tutorNotificationEmail = await readJSONValue(JSONwithTS, 'sendNotificationTutor');
                let tutorEmail = await readJSONValue(JSONwithTS, 'sendEmailTutor');
                let tutor15minEmail = await readJSONValue(JSONwithTS, 'tutorEmail15');
                let min30EmailAbsentStudent = await readJSONValue(JSONwithTS, 'studentAbsentMore30min');
                let lastEmailTime;
                let timeEmailInfo;
                if (meetingDuration > 60) {
                    lastEmailTime = after44mins;
                    timeEmailInfo = '45';
                } else {
                    lastEmailTime = after30mins;
                    timeEmailInfo = '30';
                }
                if (absentee.role === 'tutor') {
                    if (timeRightNow > after3mins && tutorEmail == 0) {
                        absenteesInfo = absentee.groups;
                        absenteesInfo.forEach(guestTutor => {
                            presentUserEmail = guestTutor.email;
                            absenteeName = guestTutor.name;
                        });
                        // Check if there are present students
                        const presentStudents = uniquePresent.filter(user => user.role === 'student');
                        if (presentStudents.length > 0) {
                            const presentStudentEmails = presentStudents.map(student => student.email).join(", ");
                            const presentStudentName = presentStudents.map(student => student.email).join(", ");
                            const notificationSlackPresentStudent = `${(new Date()).toLocaleTimeString('en-US')} \n:incoming_envelope: Notification sent to ${presentStudentName}`
                            sendingEmailPresentUser(absenteeName, presentStudentEmails, eventInfo);
                            await sendMessage(thread, notificationSlackPresentStudent);

                        }
                        await updateTsValue(JSONwithTS, 'sendEmailTutor', tutorEmail + 1);
                    }
                } else if (absentee.role === 'student' && absentee.groups.length > 0) {
                    if (
                        (timeRightNow > after5mins && studentEmail == 0 && studentFlag == 0) ||
                        (timeRightNow > after10mins && studentEmail == 1 && studentFlag == 0) ||
                        (timeRightNow > after15mins && studentEmail == 2 && studentFlag == 0)
                    ) {
                        absenteesInfo = absentee.groups;
                        let emailAddresses = [];
                        let studentNames = [];
                        absenteesInfo.forEach(guestStudent => {
                            studentNames.push(guestStudent.name);
                            if (guestStudent.parent_email) {
                                emailAddresses.push(`${guestStudent.email}, ${guestStudent.parent_email}`);
                            } else {
                                emailAddresses.push(guestStudent.email);
                            }
                        });
                        presentUserEmail = emailAddresses.join(", ");
                        absenteeName = studentNames.join(", ");
                        // Check if there are present tutors
                        const presentTutors = uniquePresent.filter(user => user.role === 'tutor');
                        if (presentTutors.length > 0 && tutorNotificationEmail == 0 && reviewIgGroupClass != true) {
                            const presentTutorEmails = presentTutors.map(tutor => tutor.email).join(", ");
                            const presentTutorNames = presentTutors.map(tutor => tutor.name).join(", ");
                            sendingEmailPresentUser(absenteeName, presentTutorEmails, eventInfo);
                            const notificationSlackPresentTutor = `${(new Date()).toLocaleTimeString('en-US')} \n:incoming_envelope: Notification sent to ${presentTutorNames}`
                            await sendMessage(thread, notificationSlackPresentTutor);
                            await updateTsValue(JSONwithTS, 'sendNotificationTutor', tutorNotificationEmail + 1);
                        }
                        await updateTsValue(JSONwithTS, 'sendEmailStudent', studentEmail + 1);
                    } else if (timeRightNow > after15mins && tutor15minEmail == 0  && studentFlag == 0) {
                        const hasPresentStudents = uniquePresent.some(user => user.role === 'student');
                        if (!hasPresentStudents) {
                            const presentTutors = uniquePresent.filter(user => user.role === 'tutor');
                            if (presentTutors.length > 0 && reviewIgGroupClass != true) {
                                const presentTutorEmails = presentTutors.map(tutor => tutor.email).join(", ");
                                const presentTutorNames = presentTutors.map(tutor => tutor.name).join(", ");
                                sendingEmailPresentTutor15min(presentTutorEmails, eventInfo);
                                const notificationSlackPresentTutor = `${(new Date()).toLocaleTimeString('en-US')} \n:incoming_envelope: 15 minutes mark notification sent to ${presentTutorNames}\n<@U05M0P95AGZ>`
                                await sendMessage(thread, notificationSlackPresentTutor);
                            }
                            await updateTsValue(JSONwithTS, 'tutorEmail15', tutor15minEmail + 1);
                        }
                    } else if (timeRightNow > lastEmailTime && min30EmailAbsentStudent == 0  && studentFlag == 0) {
                        const hasPresentStudents = uniquePresent.some(user => user.role === 'student');
                        if (!hasPresentStudents) {
                            const presentTutors = uniquePresent.filter(user => user.role === 'tutor');
                            if (presentTutors.length > 0 && reviewIgGroupClass != true) {
                                const presentTutorEmails = presentTutors.map(tutor => tutor.email).join(", ");
                                const presentTutorNames = presentTutors.map(tutor => tutor.name).join(", ");
                                sendingEmailPresentTutor30min(timeEmailInfo, presentTutorEmails, eventInfo)
                                const notificationSlackPresentTutor = `${(new Date()).toLocaleTimeString('en-US')} \n:incoming_envelope: 30/45 minutes mark notification sent to ${presentTutorNames}\n<@U05M0P95AGZ>`
                                await sendMessage(thread, notificationSlackPresentTutor);
                            }
                            await updateTsValue(JSONwithTS, 'studentAbsentMore30min', min30EmailAbsentStudent + 1);
                        }
                    }
                }

                if (presentUserEmail) {
                    sendingEmailAbsent(eventInfo, presentUserEmail, page, attendanceScreenshot);
                    let notificationSlackAbsentee = `${(new Date()).toLocaleTimeString('en-US')} \n:incoming_envelope: Reminder sent to ${absenteeName}
                            \n<@U05M0P95AGZ>`;
                    await sendMessage(thread, notificationSlackAbsentee);
                }

            }
        }
        await sleep(interval);
        return log;
    }



    async leaveSession(meetLink) {

        const page = this.page;

        await page.close();

        console.log(`tab closed! - ${meetLink} `);

        await sleep(1000);

    }



    async isInMeet() {

        const page = this.page;

        const element = await page.$('[class="P245vb"]');

        if (element) {

            return true;

        } else {

            return false;

        }

    }

}


async function removeAMorPM(message) {
    const match = message.match(/(AM|PM)(\s[\s\S]*)/);

    let text;

    if (match) {
        text = match[2].trim(); // This will contain the part after "PM" or "AM"
    } else {
        text = message
    }

    return text
}


async function monitorMeet(browser, event) {

    let eventID = event.id;
    let attendanceLog = "";
    let fileName = `Record_${eventID}`
    await writeJSONFile(fileName)
    // const startTime = event.startTime.getTime();

    const endTime = event.endTime.getTime();
 
    // create new window with new tab
    const bot = await AttendanceBot.initialize(browser);

    const patternGroup = /^\b[^\W\d_]+\b\s*:\s*(?:\s*\b[^\W\d_]+\b)+/;// Regex pattern for Group Classes

    try {

        // find or create Slack thread

        const thread = await getThread(event.title, event.startTime, event.endTime);

        const isGroupClass = patternGroup.test(event.title);

        if (isGroupClass) {

            await updateTsValue(fileName, "isGroupClass", true)

        }



        // join session

        await bot.joinSession(event.title, event.meetLink, thread);

        bot.postContactGuest(thread, event.guests)

        await sleep(1000);

        // auto admit AND take attendance

        await pAll([

            async () => {

                while (Date.now() < endTime) {

                    if (await bot.isInMeet() === true) {

                        await bot.autoAdmit(5 * 1000);

                    } else {

                        sendMessage(thread, "Removed from session 😭")

                        throw Error("bot removed :(");

                    }

                }

            },

            async () => {



                console.log("taking attendance...");

                while (Date.now() < endTime + (4 * 60 * 1000)) {

                    if (await bot.isInMeet() === true) {

                        let guestListForAttandace = event.guests.slice();
                        guestListForAttandace.push({
                            email: 'bot@noemail.com',
                            responseStatus: 'needsAction'
                        });

                        //console.log(guestListForAttandace)

                        await bot.takeAttendance(
                            thread,
                            1 * 60 * 1000,
                            guestListForAttandace,
                            event,
                            fileName
                        ); // CHANGE WHEN TESTING

                    } else {

                        throw Error("bot removed :(");

                    }

                }

            },

        ]);



        // wait until end of session plus 5 minutes

        const untilEnd = event.endTime - Date.now();

        await sleep(untilEnd + (5 * 60 * 1000)); // SAFE TO COMMENT OUT WHEN TESTING

    } catch (e) {

        console.log(e);

    } finally {

        // leave session and close tab

        await bot.leaveSession(event.meetLink);

        console.log(`closing context...${event.meetLink} `);


        deleteJSONFile(fileName)

        await bot.context.close();

        return attendanceLog;

    }

}

async function getEventData(timeMin, timeMax) {

    var calendarListRes = await calendar.calendarList.list({ maxResults: 250 });

    var calendarList = calendarListRes.data.items;

    while (calendarListRes.data.nextPageToken != null) {

        calendarListRes = await calendar.calendarList.list({ maxResults: 250, pageToken: calendarListRes.data.nextPageToken });

        calendarList = calendarList.concat(calendarListRes.data.items);

    }

    const calendarEvents = [];

    const events = [];

    for (const calendarItem of calendarList) {

        const calendarId = calendarItem.id;

        const eventListRes = await calendar.events.list({

            calendarId,

            timeMin,

            timeMax,

            singleEvents: true,

            orderBy: 'startTime',

        });

        var eventList = [];

        eventList = eventListRes.data.items;

        for (const event of eventList) {

            calendarEvents.push(event);

        }

    }


    console.log(`Retrieved ${calendarEvents.length} events from ${calendarList.length} calendars.`);



    for (const calendarEvent of calendarEvents) {

        var monitoringInvited = false;

        const attendeeList = calendarEvent.attendees; //Gets the guest list 

        if (attendeeList == null) { //Skip event if there are no attendees

            continue;

        }

        for (const attendee of attendeeList) {

            if (attendee.email == MONITORING_EMAIL) { //Check if guest email is monitoring

                monitoringInvited = true;

                break;

            }

        }

        if (!monitoringInvited || calendarEvent.summary.match(CANCELLED_REGEX) != null) { //Check if monitoring invited and if event is not cancelled

            continue;

        }

        const event = new Event(

            calendarEvent.summary,

            new Date(calendarEvent.start.dateTime),

            new Date(calendarEvent.end.dateTime),

            calendarEvent.hangoutLink,

            attendeeList,

            calendarEvent.id


        );

        events.push(event);

    }

    return events;

}



// HELPERS

/** 
 
 * saves cookies from the current execution so login credentials can be 
 
 * remembered between runs without hardcoding passwords
 
 * @param {Promise} page the current page object that cookies will be saved for
 
 * @returns {void}
 
 */

async function saveCookies(page, filePath) {

    const cookies = await page.cookies();

    await fs.writeFile(filePath, JSON.stringify(cookies, null, 2));

}



async function loadCookies(filePath) {

    const cookiesData = await fs.readFile(filePath);

    return JSON.parse(cookiesData);

}



function areCookiesValid(cookies) {

    // check if the cookies array is empty

    if (cookies.length === 0) {

        return false;

    }

    // checks if any cookie is expired

    const currentTimestamp = Date.now() / 1000;

    for (const cookie of cookies) {

        if (cookie.expires < currentTimestamp) {

            return false;

        }

    }

    return true;

}



async function requestLogin() {

    let loginInfo = [];

    const schema = {

        properties: {

            gmail: {

                description: 'please enter your Gmail username ',

                pattern: /^[a-z0-9](\.?[a-z0-9]){5,}@([\w-]*)\.com$/,

                message: 'Gmail username must end in @[domain].com',

                required: true

            },

            pass: {

                description: 'please enter your Gmail password (not saved) ',

                hidden: true

            }

        }

    };



    if (process.env.GMAIL === "" || process.env.PASSWORD === "") {

        // start the prompt

        prompt.start();



        // get login credentials from user, not hardcoded in .env file

        loginInfo = await prompt.get(schema);

        console.log('login credentials received...');

        await sleep(3000);



        // this does not save user input to the .env file, it is only stored for the current execution

        if (loginInfo != undefined) {

            process.env.GMAIL = loginInfo.gmail;

            process.env.PASSWORD = loginInfo.pass;

        }

    }

}



/** 
 
 * sleep function 
 
 * @param {int} milliseconds the number of milliseconds the script will pause for
 
 * @returns {Promise} the response of setTimeout
 
 */

function sleep(milliseconds) {

    return new Promise((r) => setTimeout(r, milliseconds));

}




// Takes list of displayed guests and list of invited emails; 
//returns list of Absentees and the present users with their database information

async function getGuestUsers(invitees) {
    let guestUsers = [];

    const PEOPLEDATABASE = await getGoogleSheetsData(spreadsheetId, ALLTABNAME);

    for (const invite of invitees) {
        if (invite.email == OPERATION_EMAIL || invite.email == MONITORING_EMAIL) {
            continue; // Skip checking monitoring email
        }

        let invitedName;
        let invitedNickName;
        let role;
        let guestEmail = invite.email;

        if (guestEmail == 'bot@noemail.com') {
            invitedName = firefliesAccount;
            invitedNickName = 'Fireflies.ai Notetaker';
            role = 'bot';
            guestUsers.push({ name: invitedName, nickname: invitedNickName, role: role, email: guestEmail });

        } else {
            let peopleData = getPersonFromData(PEOPLEDATABASE, guestEmail, EMAIL_COLUMN);
            for (let personData of peopleData) {
                if (personData) {
                    invitedName = personData[NAME_COLUMN - 1] || 'No name';
                    invitedNickName = personData[NICKNAME_COLUMN - 1] || 'No nickname';
                    role = personData[ROLE_COLUMN - 1] || 'No role';

                } else {
                    // Convert email to string to prevent Slack mailto link
                    invitedName = `<mailto:${guestEmail}|${guestEmail}>`;
                    invitedNickName = 'Not in DB';
                    role = 'Not in DB';
                }
                if (invitedName != undefined) {
                    guestUsers.push({ name: invitedName, nickname: invitedNickName, role: role, email: guestEmail });
                }
            }
        }
    }

    return guestUsers;
}



function getPresentUsers(guestUsers, displayList) {
    let present = [];

    //console.dir(guestUsers)

    guestUsers.forEach((guestUser) => {
        const nameToSearchInvated = guestUser.name.toLowerCase();
        const nicknameToSearchInvated = guestUser.nickname ? guestUser.nickname.toLowerCase() : null;

        const foundDisplayUser = displayList.find(
            (displayUser) =>
                displayUser.name.toLowerCase() == nameToSearchInvated ||
                (nicknameToSearchInvated && displayUser.name.toLowerCase() == nicknameToSearchInvated)
        );

        if (foundDisplayUser) {
            present.push({
                name: guestUser.name,
                nickname: guestUser.nickname,
                role: guestUser.role,
                email: guestUser.email
            });
        }
    });

    // Add users from displayList that are not in guestUsers
    displayList.forEach((displayUser) => {
        const nameToSearch = displayUser.name.toLowerCase();
        const nicknameToSearch = displayUser.nickname ? displayUser.nickname.toLowerCase() : null;

        const isUserAlreadyAdded = present.some(
            (addedUser) =>
                addedUser.name.toLowerCase() == nameToSearch ||
                (nicknameToSearch && addedUser.name.toLowerCase() == nicknameToSearch)
        );

        if (!isUserAlreadyAdded) {
            if (displayUser.name == 'ModernSmart Operation') { return }
            else if (displayUser.name.includes('Fireflies.ai Notetaker')) {
                present.push({
                    name: firefliesAccount,
                    role: 'Bot',
                    email: 'bot@noemail.com'
                });

            }
        } else {
            // Add the user to the present list
            present.push({
                name: displayUser.name,
                nickname: 'no nickname',
                role: 'Not found in DB',
            });
        }
    });

    return present;
}


function getMissingUsers(guestUsers, displayList) {
    // Make a copy of the guestUsers list
    let guestUsersList = [...guestUsers];

    let guestUsersNew = removeDuplicateTutorAndEmail(guestUsersList);

    displayList.forEach(displayUser => {
        if (displayUser.name.includes('Fireflies.ai Notetaker')) {
            // Find the index of the matching user in guestUsersNew
            const index = guestUsersNew.findIndex(user => (
                user.name.toLowerCase().includes('fireflies.ai notetaker') ||
                (user.nickname && user.nickname.toLowerCase().includes('fireflies.ai notetaker'))
            ));

            // If the user is found, remove it from guestUsersNew
            if (index !== -1) {
                guestUsersNew.splice(index, 1);
            }
        }
    });

    // Array to store indexes to be removed
    let indexesToRemove = [];

    // Loop through each object in displayList
    displayList.forEach((displayUser, displayIndex) => {
        // Find and remove matching user from guestUsersNew
        let foundGuestIndex = guestUsersNew.findIndex(guestUser =>
            guestUser.name.toLowerCase() == displayUser.name.toLowerCase() ||
            (guestUser.nickname && guestUser.nickname.toLowerCase() == displayUser.name.toLowerCase())
        );

        //console.log(`Display Index: ${displayIndex}, foundGuestIndex: ${foundGuestIndex}`);

        if (foundGuestIndex !== -1) {
            indexesToRemove.push(foundGuestIndex);

            // Remove the foundGuest from guestUsersNew if name or nickname matches
            let remainingGuestIndexes = guestUsersNew.reduce((acc, guestUser, index) => {
                if (
                    index !== foundGuestIndex &&
                    ((guestUser.name.toLowerCase() == guestUsersNew[foundGuestIndex].name.toLowerCase()) ||
                        (guestUser.nickname != 'No nickname' && guestUser.nickname.toLowerCase() == guestUsersNew[foundGuestIndex].name.toLowerCase()) ||
                        (guestUser.nickname != 'No nickname' && guestUser.nickname.toLowerCase() == guestUsersNew[foundGuestIndex].nickname.toLowerCase()))
                ) {
                    acc.push(index);
                }
                return acc;
            }, []);

            //console.log(`Remaining Indexes: ${remainingGuestIndexes}`);

            // Add all remainingGuestIndexes to indexesToRemove
            indexesToRemove = [...indexesToRemove, ...remainingGuestIndexes];
        }
    });

    // Remove duplicates from indexesToRemove
    indexesToRemove = [...new Set(indexesToRemove)];

    // Sort indexesToRemove in descending order to avoid issues when splicing
    indexesToRemove.sort((a, b) => b - a);

    // Remove elements from guestUsersNew based on indexesToRemove
    indexesToRemove.forEach(index => {
        guestUsersNew.splice(index, 1);
    });

    // The remaining objects in guestUsersNew are the missing users
    let missing = guestUsersNew;

    //console.log('missing: before filtering duplicates');
    //console.dir(missing);

    return missing;
}


function removeDuplicateTutorAndEmail(guestObjectsList) {
    const tutorEmailToRemove = process.env.PRINCETON_EMAIL;

    // Check if there is any other object with the role Tutor
    const hasOtherTutor = guestObjectsList.some(obj => obj.role.toLowerCase() === 'tutor');

    // Check if there is an object with the email address to remove
    const hasEmailToRemove = guestObjectsList.some(obj => obj.email === tutorEmailToRemove);

    // If both conditions are met, remove the object with the specified email
    if (hasOtherTutor && hasEmailToRemove) {
        return guestObjectsList.filter(obj => obj.email !== tutorEmailToRemove);
    }

    return guestObjectsList;
}

async function createListGuest(guestDisplayList, guestList) {

    let display = guestDisplayList;

    let invitees = guestList;


    const guestUsers = await getGuestUsers(invitees);

    const present = getPresentUsers(guestUsers, display);
    const missing = getMissingUsers(guestUsers, display);

    present.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
    missing.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

    // Create a map to track unique names and nicknames
    const uniqueNamesAndNicknames = new Map();

    // Filter duplicates in present
    const uniquePresent = present.filter((value) => {
        const key = value.name.toLowerCase();
        if (!uniqueNamesAndNicknames.has(key)) {
            uniqueNamesAndNicknames.set(key, true);
            return true;
        }
        return false;
    });

    // Clear the map for re-use
    uniqueNamesAndNicknames.clear();

    // Filter duplicates in missing
    const uniqueMissing = missing.filter((value) => {
        const key = value.name.toLowerCase();
        if (!uniqueNamesAndNicknames.has(key)) {
            uniqueNamesAndNicknames.set(key, true);
            return true;
        }
        return false;
    });


    return { uniquePresent, uniqueMissing };
}


// Called on each meeting every 1 minutes for the first ~20 minutes

async function createLog(guests, absentees) {
    let log = "";

    //console.log('createLog guest ')

    //console.dir(guests)
    //console.log('createLog absenttes')
    //console.dir(absentees)

    log += (new Date()).toLocaleTimeString('en-US');

    // Log present guests
    if (guests && guests.length != 0) {
        log += "\nPresent:\n";
        for (const guest of guests) {
            //console.dir(guest);
            log += `∙ ${guest.name} :white_check_mark: ${capitalizeFirstLetter(guest.role)} \n`;
        }
    }

    // Log absentees
    if (absentees && absentees.length != 0) {
        let logabs = "";
        for (const absentee of absentees) {
            if (absentee.role != "parent") { // Exclude the parents since the contact information will be posted at the beginning
                logabs += `∙ ${absentee.name} :x: ${capitalizeFirstLetter(absentee.role)}\n`;
            }
        }

        if (logabs.trim() != '') {
            log += `\nAbsent: \n${logabs}`;
        }
    }
    //console.dir(log)
    return log;
}



async function notifyAbsence(absentee, eventInfoData, thread, JSONwithTS, keyTutor, keyStudent, keyBot) {
    try {
        let meetingTime = eventInfoData.startTime
        let titleAbsenceRecord = JSONwithTS;
        const curTime = new Date();
        const lateAmount = curTime - meetingTime;
        const role = absentee.role;
        //const name = absentee.name;
        const newMeetingTime = new Date(meetingTime)
        const after4mins = new Date(newMeetingTime.getTime() + (4 * 60 + 40) * 1000);
        const after6mins = new Date(newMeetingTime.getTime() + (5 * 60 + 40) * 1000);
        const after7mins = new Date(newMeetingTime.getTime() + (7 * 60 + 40) * 1000);
        const after9mins = new Date(newMeetingTime.getTime() + (8 * 60 + 40) * 1000);
        let lateThreshold;
        let keyTS;

        // console.dir('before the log: ', absentee)

        if (role == "student") {
            lateThreshold = STUDENTTIME;
            keyTS = keyStudent
        }

        else if (role == "tutor") {
            lateThreshold = TUTORTIME;
            keyTS = keyTutor
        }

        else if (role == "bot") {
            lateThreshold = BOTTIME;
            keyTS = keyBot;
        }
        else if (role == "parent") {
            return [];
        }

        //console.log('lateAmount:', lateAmount)
        //console.log('lateThreshold:', lateThreshold)
        let notificationLog = [];


        if (lateAmount > lateThreshold) { // Person is absent past the allowed time amount
            //console.log(titleAbsenceRecord);

            if (absentee.groups && absentee.groups.length > 0) {
                let groupedUsers = absentee.groups;
                groupedUsers.forEach(userAbsent => {
                    if (userAbsent.name != '') {
                        let userNAmeGroup = `:bangbang: ${userAbsent.name} is absent.\n`
                        notificationLog.push(userNAmeGroup)
                    }
                });
            } else {

                notificationLog = [`:bangbang: ${absentee.name} is absent.\n `];

            }

            let contentAbsent = `${curTime.toLocaleTimeString('en-US')}\n${notificationLog.join('\n')}\n<@U05M0P95AGZ>`;
            console.log('log to be post: ', contentAbsent)
            let replyAbsentID = await readJSONValue(titleAbsenceRecord, keyTS);
            let currentabsentMessage;

            if (replyAbsentID != "0") {
                currentabsentMessage = await getMessageByTsAndChannel(MONITORING_CHANNEL_ID, replyAbsentID);
            } else {
                currentabsentMessage = 'PM';
            }

            let recordAbsent1 = await removeAMorPM(contentAbsent);

            if (
                (role === "tutor" && (after4mins < curTime && curTime < after6mins || after7mins < curTime && curTime < after9mins)) ||
                (role === "student" && after7mins < curTime && curTime < after9mins)
            ) {
                recordAbsent1 = '';
            }




            let recordAbsent2 = await removeAMorPM(currentabsentMessage);

            if (recordAbsent1 !== recordAbsent2) {
                await sendMessagewithrecord(thread, contentAbsent, keyTS, titleAbsenceRecord);
            } else {
                // The string representations are the same, so no need to send a message.
                console.log("There is no change in absent list");
            }

            return [absentee.email]; // Mark this absentee as late, so that they are only notified once.
        }



        return [];
    } catch (error) {
        console.error("Error in notifyAbsence:", error);
        return [];
    }
}



// //Given two names (display name and name from database), assess if they are equal

// function matchNames(displayName, databaseName) {

//     const charRegex = /[^a-zA-Z]/g;

//     var name1 = displayName.toLowerCase().trim(); //Ignore casing and leading/trailing whitespace

//     var name2 = databaseName.toLowerCase().trim();

//     // Reduce name to only characters, then sort to compare anagrams, which accounts for rearranged first/last/middle names

//     var name1chars = name1.replace(charRegex, '').split("").sort().join("");

//     var name2chars = name2.replace(charRegex, '').split("").sort().join("");

//     if (name1chars == name2chars) { return true; }

//     var name1words = name1.split(" ");

//     var name2words = name2.split(" ");

//     if (name1words[0] == name2words[0]) { return true; } // Return true if first names match

//     if (name1words[name1words.length - 1] == name2words[name2words.length - 1]) { return true; } // Return true if last names match

//     return false;

// }



// Gets data from a Google Sheets tab (assuming credentials allow for access)

// Input: (String, String); ID of Google Sheets, name of tab in the sheet

// Output: (Array[String[]]): Ordered array of all rows in the tab

async function getGoogleSheetsData(spreadsheetId, tabName) {
    let response = "";
    try {
        const auth = await authorize(); // Obtain OAuth2Client instance
        const sheetsAPI = google.sheets({ version: "v4", auth });

        response = await sheetsAPI.spreadsheets.values.get({
            spreadsheetId,
            range: `${tabName}!A2:I`, // Get all rows except the first (column names)
        });

        console.log(`Loaded ${response.data.values.length} rows from Google Sheets.`)
    } catch (e) {
        console.error("Error loading Google Sheets data:", e);
        return []; // Return empty array if there is an error
    }

    return response.data.values;
}




// Given an identifier, searches the database and returns their data

// Input: (Array[String[]], String, Integer); Database, Identifier, column number of identifier (starts at 1)

// Output: (String[]); allthe rows of data matching the identifier

function getPersonFromData(database, identifier, columnNum) {
    const matchingRows = [];

    try {

        for (let i = 0; i < database.length; i++) {
            const row = database[i];

            if (
                row[columnNum - 1]?.toLowerCase().includes(identifier.toLowerCase()) &&
                row[ROLE_COLUMN - 1]?.toLowerCase() !== 'staff' &&
                row[ROLE_COLUMN - 1]?.toLowerCase() !== 'partner'
            ) {
                // Found a matching row
                matchingRows.push(row); // Push the rowData array directly
            }
        }
        return matchingRows;
    } catch (e) {
        console.error(e)
    }


}



async function getThread(title, startTime, endTime) {

    let startString = startTime.toLocaleTimeString("en-US");

    let endString = endTime.toLocaleTimeString("en-US");

    let date = startTime.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });

    startString = (startString.match(DATE_REGEX))[1] + (startString.match(DATE_REGEX))[2].toLowerCase();

    endString = (endString.match(DATE_REGEX))[1] + (endString.match(DATE_REGEX))[2].toLowerCase();

    const text = `${title} ${date}⋅${startString} - ${endString}`

    // console.log(title, date, startString, endString)

    try {

        const thread = await findMessage(MONITORING_CHANNEL_ID, [title, date, startString, endString]);

        console.log("existing thread found!");

        return thread;

    } catch (e) {

        console.log(e)

        console.log("creating new thread.");

        const thread = await publishMessage(MONITORING_CHANNEL_ID, text);

        return thread;

    }

}



function sendMessage(thread, message) {

    mutex.runExclusive(async () => {

        try {

            await replyMessage(MONITORING_CHANNEL_ID, thread.ts, message);

            console.log("message sent!");

        } catch (e) {

            console.log(e);

        }

    });

}

function sendMessagewithrecord(thread, message, keyValue, nameRecord) {

    let title = nameRecord;
    let keyTSValue = keyValue;

    mutex.runExclusive(async () => {

        try {

            await replyMessagewithrecord(MONITORING_CHANNEL_ID, thread.ts, message, keyTSValue, title);

            console.log("message sent!");

        } catch (e) {

            console.log(e);

        }

    });

}

async function sendEmail(auth, to, bcc, subject, body, attachmentBase64 = null) {
    const gmail = google.gmail({ version: 'v1', auth });

    console.log('Sending email')

    // Compose email message
    let message = `To: ${to}\r\nBCC: ${bcc}\r\n` +
        `Subject: ${subject}\r\n` +
        'Content-Type: multipart/mixed; boundary="boundary-example"\r\n' +
        '\r\n' +
        '--boundary-example\r\n' +
        'Content-Type: text/html; charset="UTF-8"\r\n' +
        '\r\n' +
        `${body}\r\n`;

    if (attachmentBase64) {
        message += '--boundary-example\r\n' +
            'Content-Type: image/png; name="attachment.png"\r\n' +
            'Content-Disposition: attachment; filename="attachment.png"\r\n' +
            'Content-Transfer-Encoding: base64\r\n' +
            '\r\n' +
            `${attachmentBase64}\r\n`;
    }

    message += '--boundary-example--';

    console.log('Email Message:', message);

    const encodedMessage = Buffer.from(message).toString('base64');

    // Send email
    gmail.users.messages.send({
        userId: 'me',
        resource: {
            raw: encodedMessage,
        },
    }, (err, res) => {
        if (err) return console.error('The API returned an error:', err.message);
        console.log('Email sent:', res.data);
    });
}

async function sendingEmailAbsent(event, email, page, pathName) {
    try {
        const auth = await authorize();
        let dtStart = DateTime.fromISO(event.startTime.toISOString(), { zone: 'America/New_York' });
        let dtEnd = DateTime.fromISO(event.endTime.toISOString(), { zone: 'America/New_York' });
        let eventStartTime = dtStart.toFormat("EEEE LLL dd, yyyy · h:mm a");
        let eventEndTime = dtEnd.toFormat("h:mm a");
        let eventTitle = event.title.trim();
        const toAddress = process.env.NOTIFICATION_EMAIL;
        const bccAddress = email;
        let imagePathNameAttendance = `./record/${pathName}_attachments.png`;
        await page.waitForTimeout(8000);
        let attachmentPic = await page.screenshot({ path: imagePathNameAttendance, fullPage: true });

        const emailSubject = `Please Join Now: ${eventTitle} (ModernSmart Team)`;
        const emailBody = `
        <body style="font-family: 'Roboto', sans-serif;">
            <p> Reminder </p>
            <table style="border-collapse: collapse; border: 2px solid lightgray; border-radius: 10px;">
                <tr>
                    <td style="padding: 10px;">
                        <h2>${eventTitle}</h2>
                        <p>${eventStartTime} - ${eventEndTime} (Eastern time - New York)</p>
                        <h3>Meeting link</h3>
                        <p><a href="${event.meetLink}">${event.meetLink}</a></p>
                    </td>
                </tr>
            </table>
            <br>
            <p style="margin-bottom: 2px;"> This is an automated message. If you already notified us of your absence, please ignore this message. </p>
            <p style="margin-bottom: 2px;">This message is for the session ${eventTitle} @ ${eventStartTime} - ${eventEndTime}</p>
        </body>
        `;

        await sendEmail(auth, toAddress, bccAddress, emailSubject, emailBody, attachmentPic.toString('base64'));
        await fs.unlink(imagePathNameAttendance);
    } catch (error) {
        console.error(error);
    }
}


async function sendingEmailPresentUser(absenteeName, email, event) {
    try {
        const auth = await authorize();
        let dtStart = DateTime.fromISO(event.startTime.toISOString(), { zone: 'America/New_York' });
        let dtEnd = DateTime.fromISO(event.endTime.toISOString(), { zone: 'America/New_York' });
        let eventStartTime = dtStart.toFormat("EEEE LLL dd, yyyy · h:mm a");
        let eventEndTime = dtEnd.toFormat("h:mm a");
        let eventTitle = event.title.trim();
        const toAddress = process.env.NOTIFICATION_EMAIL;
        const bccAddress = email;
        const emailSubject = `Quick Update: Contacting ${absenteeName} to join`;
        const emailBody = `
        <body style="font-family: 'Roboto', sans-serif;">
            <p>Hello!</p>
            <p>Just a quick heads up - We are currently contacting ${absenteeName} to have them join the class.</p>
            <p>Thanks for your patience!</p>
            <p style="margin-bottom: 2px;">Best regards,<br>ModernSmart Operation</p>
            <p style="margin-bottom: 2px;">This message is for the session ${eventTitle} @ ${eventStartTime} - ${eventEndTime}</p>
        </body>
    `;

        await sendEmail(auth, toAddress, bccAddress, emailSubject, emailBody, null);

    } catch (error) {
        console.error(error);
    }
}

async function sendingEmailPresentTutor15min(email, event) {
    try {
        const auth = await authorize();
        let dtStart = DateTime.fromISO(event.startTime.toISOString(), { zone: 'America/New_York' });
        let dtEnd = DateTime.fromISO(event.endTime.toISOString(), { zone: 'America/New_York' });
        let eventStartTime = dtStart.toFormat("EEEE LLL dd, yyyy · h:mm a");
        let eventEndTime = dtEnd.toFormat("h:mm a");
        let eventTitle = event.title.trim();
        const toAddress = process.env.NOTIFICATION_EMAIL;
        const bccAddress = email;
        const emailSubject = `Quick Update: 15-minute update`;
        const emailBody = `
        <body style="font-family: 'Roboto', sans-serif; max-width: 600px;">
        
          <p style="margin-bottom: 2px;">Hello!</p>
        
          <p style="margin-bottom: 2px;">We have been continuously trying to contact the student to ensure their participation in the session, but they have not yet been able to join up to now. If it's a one-hour session, you may wait up to 30 minutes before leaving, and for sessions longer than an hour, you may wait up to 45 minutes. From now on, you can turn off your camera and only leave your speaker on while waiting for the student. If you have any questions/comments, you can respond to this email. Thank you.</p>
        
          <p style="margin-bottom: 2px;">Best regards,<br>
          ModernSmart Operation</p>
          <p style="margin-bottom: 2px;">This message is for the session ${eventTitle} @ ${eventStartTime} - ${eventEndTime}</p>
        
        </body>`;

        await sendEmail(auth, toAddress, bccAddress, emailSubject, emailBody, null);

    } catch (error) {
        console.error(error);
    }
}


async function sendingEmailPresentTutor30min(time, email, event) {
    try {
        const auth = await authorize();
        let dtStart = DateTime.fromISO(event.startTime.toISOString(), { zone: 'America/New_York' });
        let dtEnd = DateTime.fromISO(event.endTime.toISOString(), { zone: 'America/New_York' });
        let eventStartTime = dtStart.toFormat("EEEE LLL dd, yyyy · h:mm a");
        let eventEndTime = dtEnd.toFormat("h:mm a");
        let eventTitle = event.title.trim();
        const toAddress = process.env.NOTIFICATION_EMAIL;
        const bccAddress = email;
        const emailSubject = `Quick Update: ${time}-minute update`;
        const emailBody = `
        <body style="font-family: 'Roboto', sans-serif; max-width: 600px;">
        
          <p style="margin-bottom: 2px;">Hello!</p>
        
          <p style="margin-bottom: 2px;">This is an automated message. Unless you have been notified by us or the student otherwise, please feel free to exit the session when more than ${time} minutes have elapsed. You will be compensated at the full rate for this no-show session. If you have any questions/comments, you can respond to this email. Thank you.</p>
        
          <p style="margin-bottom: 2px;">Best regards,<br>
          ModernSmart Operation</p>
          <p style="margin-bottom: 2px;">This message is for the session ${eventTitle} @ ${eventStartTime} - ${eventEndTime}</p>
        
        </body>`;

        await sendEmail(auth, toAddress, bccAddress, emailSubject, emailBody, null);

    } catch (error) {
        console.error(error);
    }
}


async function createBrowser() {

    // create browser dummy

    const option = {

        revision: "",

        detectionPath: "",

        folderName: ".chromium-browser-snapshots",

        defaultHosts: ["https://storage.googleapis.com", "https://npm.taobao.org/mirrors"],

        hosts: [],

        cacheRevisions: 2,

        retry: 3,

        silent: false

    };

    const stats = await PCR(option);



    const browser = await puppeteerExtra.launch({
        headless: "new", // Opt-in for the new Headless mode
        executablePath: stats.executablePath,
        args: [
            '--no-sandbox',
            '--use-fake-ui-for-media-stream' // Note: Removed the trailing comma
        ]
    });
    return browser;

}



async function authorize() {

    const credentials = require(CLIENT_SECRET_PATH);



    const oAuth2Client = new google.auth.OAuth2(

        credentials.installed.client_id,

        credentials.installed.client_secret,

        credentials.installed.redirect_uris[0]

    );



    try {

        const token = require(`./${TOKEN_PATH}`);

        oAuth2Client.setCredentials(token);

        return oAuth2Client;

    } catch (err) {

        console.log("Token missing or incorrect.")

    }

}

// Function to create a JSON file, with the event ID as title

async function writeJSONFile(fileName) {
    const __dirname = path.dirname(fileURLToPath(import.meta.url));

    const jsonFilePath = path.join(__dirname, 'record', `${fileName}.json`);
    const data = {
        absenceTs: "0",
        isGroupClass: false,
        attendaceTuTorTs: "0",
        attendaceStudentTs: "0",
        attendaceBotTs: "0",
        missingStudentFlag: 0,
        sendEmailTutor: 0,
        sendEmailStudent: 0,
        sendNotificationTutor: 0,
        postscrenshot: 0,
        tutorEmail15: 0,
        studentAbsentMore30min: 0
    };
    const jsonData = JSON.stringify(data, null, 2);
    console.log("File created")

    try {
        await fs.writeFile(jsonFilePath, jsonData, 'utf8');
        console.log('JSON file created');
    } catch (error) {
        console.error(error);
    }
}

// Function to write in the JSON and update the keyvalue we need 
// It will write the ts of the last post in Slack
function updateTsValue(fileName, key, newTS) {
    try {
        const jsonFilePath = `./Record/${fileName}.json`;
        const jsonData = require(jsonFilePath);

        // Update the specified key with the new value
        jsonData[key] = newTS;

        // Write the updated JSON back to the file
        const fs = require('fs');
        fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2), 'utf8');
        //console.log(`Updated "${key}" in "${fileName}.json"`);
    } catch (error) {
        console.error('Error updating JSON file:', error);
    }
}


// Function to write in the JSON file with the even name as title
// It will write the ts of the last post in Slack

async function readJSONValue(fileName, key) {
    try {
        const jsonData = require(`./Record/${fileName}.json`);
        return jsonData[key];
    } catch (error) {
        console.error('Error reading JSON file:', error);
    }
}

// function to delete the JSON record when the meets end

async function deleteJSONFile(fileName) {
    const __dirname = path.dirname(fileURLToPath(import.meta.url));

    const jsonFilePath = path.join(__dirname, 'record', `${fileName}.json`);

    try {
        await fs.unlink(jsonFilePath);
        console.log('JSON file deleted successfully');
    } catch (error) {
        console.error(error);
    }
}




function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}


export async function main() {
    try {

        // asks for login info if .env is not populated

        await requestLogin();

        // creates browser instance used by contexts

        const browser = await createBrowser();




        // interval in which event data is being gathered

        const interval = 1 * 60 * 60 * 1000; // 1 hour

        const startRange = Date.now() + (1 * 1000);

        // change to desired end date

        const endRange = (new Date("1/01/2025, 12:00:00 AM")).getTime();



        // creates triggers for when events are pulled from monitoring calendar

        const execTriggers = [];

        let timeMin = startRange;

        let timeMax = Math.min(timeMin + interval - 1, endRange - 1);



        while (timeMax < endRange) {

            const minDate = new Date(timeMin);

            const maxDate = new Date(timeMax);

            execTriggers.push(() => {
                schedule.scheduleJob(minDate, async () => {

                    // gets events from calendar

                    const events = await getEventData(minDate, maxDate); // COMMENT OUT WHEN USING LOCAL TEST CASES

                    console.log(minDate.toLocaleString('en-US'), maxDate.toLocaleString('en-US'));



                    // creates triggers for each session to be monitored

                    const botTriggers = [];

                    const size = events.length;

                    for (let i = 0; i < size; i++) {

                        const startTime = events[i].startTime;

                        if (startTime > minDate && startTime < maxDate) {

                            console.log(events[i].title);

                            // execute trigger 15 minutes before session start time OR immediately (greater of the two)

                            const timeBefore = Math.max(startTime.getTime() - (15 * 60 * 1000), Date.now() + (1 * 1000));

                            const dateBefore = new Date(timeBefore);



                            botTriggers.push(() => {
                                schedule.scheduleJob(dateBefore, async () => {

                                    const attendanceLog = await monitorMeet(

                                        browser,

                                        events[i],

                                    );

                                })
                            });

                        }

                    }

                    console.log(botTriggers.length);

                    try {

                        await pAll(botTriggers, { stopOnError: false });

                    } catch (e) {

                        console.log(e);

                    }

                })
            });

            timeMin = timeMax + 1;

            timeMax = Math.min(timeMin + interval - 1, endRange);

        }

        console.log(execTriggers.length);

        try {

            await pAll(execTriggers, { stopOnError: false });

        } catch (e) {

            console.log(e);

        }
    } catch (error) {

        console.error("An error occurred:", error);
        let lastErrorText = `:warning: There was an error in the code: 
    \n${error} \n:sos:<@U05M0P95AGZ> 
    \n:pray: Please restart the code :saluting_face:
    Channel Monitoring`;
        publishMessage(BOT_ALERT, lastErrorText)
    }

}

export { updateTsValue }

main();
