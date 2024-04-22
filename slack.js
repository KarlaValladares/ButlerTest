import { createRequire } from "module";
import { updateTsValue } from "./index.js";
const require = createRequire(import.meta.url);
const { App, LogLevel } = require("@slack/bolt");
require('dotenv').config();
import { createReadStream } from 'fs';



const token = process.env.SLACK_TOKEN;
const signingSecret = process.env.SLACK_SIGNING_SECRET;

const app = new App({
  token,
  signingSecret,
  // LogLevel can be imported and used to make debugging simpler
  // logLevel: LogLevel.DEBUG
});

async function findConversation(name) {
  try {
    // Call the conversations.list method using the built-in WebClient
    const result = await app.client.conversations.list();

    for (const channel of result.channels) {
      console.log(channel.name)
      if (channel.name === name) {
        conversationId = channel.id;

        // Print result
        console.log("Found conversation ID: " + conversationId);
        // Break from for loop
        break;
      }
    }
  }
  catch (error) {
    console.error(error);
  }
}

async function publishMessage(id, text) {
  try {
    // Call the chat.postMessage method using the built-in WebClient
    const result = await app.client.chat.postMessage({
      // The token you used to initialize your app
      channel: id,
      text: text
      // You could also use a blocks[] array to send richer content
    });

    // Print result, which includes information about the message (like TS)
    // console.log(result.response_metadata);
    return result;
  }
  catch (error) {
    console.error(error);
  }
}

async function replyMessage(id, ts, text) {
  try {
    // Call the chat.postMessage method using the built-in WebClient
    const result = await app.client.chat.postMessage({
      // The token you used to initialize your app
      channel: id,
      thread_ts: ts,
      text,
    });

    // Print result
    // console.log(result);
  }
  catch (error) {
    console.error(error);
  }
}

async function uploadImage(id, imagePath, ts) {
  const filePath = imagePath;
  const channelId = id;
  const ThreadID = ts;

  console.log('start of uploadImage()')
  console.dir(ThreadID.ts)
  console.log(`Trying to post in thread the image ${filePath}`)

  try {
    const result = await app.client.files.uploadV2({
      channel_id: channelId,
      initial_comment: "Class screenshot:",
      file: filePath,
      filename: 'class_screenshot.png', // Replace with the correct file extension
      thread_ts: ThreadID.ts,
      timeout: 10000,
    });

    console.log(result);
  } catch (error) {
    console.error(error);
  }
}


async function replyMessagewithrecord(id, ts, text, keyValue, title) {
  let nameRecord = title;
  let keyTsToUpdate = keyValue;

  console.log("reply to trhead: ", nameRecord, keyTsToUpdate)
  try {
    // Call the chat.postMessage method using the built-in WebClient
    const result = await app.client.chat.postMessage({
      // The token you used to initialize your app
      channel: id,
      thread_ts: ts,
      text,
    });

    let content = result.ts;

    await updateTsValue(nameRecord, keyTsToUpdate, content);

    return content
  }
  catch (error) {
    console.error(error);
  }
}

async function getConversationHistory(id) {
  const todayDateString = (new Date()).toLocaleDateString('en-US');

  // Store conversation history
  let conversationHistory;
  // ID of channel you watch to fetch the history for
  try {
    // Call the conversations.history method using WebClient
    const result = await app.client.conversations.history({
      channel: id,
      // limit: 20,
      oldest: (new Date(todayDateString + " 1:00:00 AM")).getTime() / 1000,
    });
    conversationHistory = result.messages;

    // Print results
    return conversationHistory.reverse();
  } catch (error) {
    console.error(error);
  }
}

async function deleteMessage(id, ts) {
  try {
    const result = app.client.chat.delete({
      channel: id,
      ts,
    })
  } catch (e) {
    console.error(e);
  }
}

async function findMessage(id, patterns) {
  const history = await getConversationHistory(id);
  // const message = history.find(msg => msg.text.includes(pattern));
  const message = history.find(msg => {
    return patterns.every(pattern => msg.text.includes(pattern));
  });

  if (message) {
    return message;
  } else {
    throw new Error("message not found.");
  }
}

async function findMessages(id, pattern) {
  const messages = []
  const history = await getConversationHistory(id);
  for (const message of history) {
    if (message.text === pattern) {
      messages.push(message);
    }
  }
  return messages;
}


async function getMessageByTsAndChannel(prod_channel_ID, messageTs) {
  try {
    // Call the conversations.history method using WebClient
    const result = await app.client.conversations.replies({
      channel: prod_channel_ID,
      ts: messageTs, // Set the latest parameter to the specific message's timestamp

    });

    if (result.messages.length > 0) {
      // Retrieve the message
      const message = result.messages[0].text;
      console.log(message)
      return message;
    } else {
      throw new Error("Message not found.");
    }
  } catch (error) {
    console.error(error);
  }
}




export {
  publishMessage,
  replyMessage,
  findMessage,
  replyMessagewithrecord,
  getMessageByTsAndChannel,
  uploadImage
};