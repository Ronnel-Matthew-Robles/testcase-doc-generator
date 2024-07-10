import axios from 'axios';
import dotenv from 'dotenv';
import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import readline from 'readline';
import OpenAI from 'openai';
import mongoose from 'mongoose';
import JiraTicketHelper from './models/JiraTicketHelper.js';
import QATestGenerator from './models/QATestGenerator.js';
import FileReference from './models/FileReference.js';
import TestCaseDocGenerator from './generate_excel_file.js';

import excel from 'exceljs';
import { title } from 'process';

// Derive __dirname in ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config();

// Connect to MongoDB
mongoose.connect('mongodb://localhost:27017');

const openai = new OpenAI({apiKey: process.env['OPENAI_API_KEY']});
const xray_token = process.env['XRAY_TOKEN'];

const jiraTicketHelperAssistantId = "asst_4Vl3xpvP7T1BRFzypol7rs8Y";
const qaTestCaseGeneratorAssistantId = "asst_AHfgXfV1JN1v6qU57h5I6dXc";

const jiraURL = "https://concentrixdev.atlassian.net";
const email = process.env["JIRA_EMAIL"];
const apiToken = process.env["JIRA_API_TOKEN"];

const xrayApiBase = "https://xray.cloud.getxray.app/api/v2";

const issueTypes = {
  TEST: 13305,
  TEST_SET: 13306,
  TEST_PLAN: 13307,
  TEST_EXECUTION: 13308,
  PRECONDITION: 13309
};

const fetchJiraUserStories = async () => {
  const url = jiraURL + "/rest/api/3/search";
  const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');
  const headers = {
    "Accept": "application/json",
    "Authorization": `Basic ${auth}`
  };

  const query = {
    'jql': 'project = WCX AND issuetype = Story AND sprint in openSprints() AND status NOT IN (Closed, Cancelled, "In Testing") ORDER BY created DESC'
  };

  const response = await axios.get(url, { headers, params: query });
  return response.data;
};

const fetchIssueDetailsWithAttachments = async (issueId) => {
  const url = `${jiraURL}/rest/api/3/issue/${issueId}`;
  const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');
  const headers = {
    "Accept": "application/json",
    "Authorization": `Basic ${auth}`
  };

  const response = await axios.get(url, { headers });
  return response.data;
};

const extractTextRecursive = (contentItem) => {
  let textParts = [];
  if (typeof contentItem === 'object' && !Array.isArray(contentItem)) {
    if ("text" in contentItem) {
      textParts.push(contentItem["text"]);
    } else if ("content" in contentItem) {
      contentItem["content"].forEach(nestedContent => {
        textParts = textParts.concat(extractTextRecursive(nestedContent));
      });
    }
  } else if (Array.isArray(contentItem)) {
    contentItem.forEach(item => {
      textParts = textParts.concat(extractTextRecursive(item));
    });
  }
  return textParts;
};

const extractText = (description) => {
  let textParts = [];
  if ("content" in description) {
    description["content"].forEach(contentItem => {
      textParts = textParts.concat(extractTextRecursive(contentItem));
    });
  }
  return textParts.join(" ");
};

// Define a function to download the image and save it locally
async function downloadImage(url, filename, headers) {
  const response = await axios({
    url,
    method: 'GET',
    responseType: 'stream',
    headers,
  });

  const localFilePath = path.resolve(__dirname, 'downloads', filename);
  const writer = fs.createWriteStream(localFilePath);

  response.data.pipe(writer);

  return new Promise((resolve, reject) => {
    writer.on('finish', () => resolve(localFilePath));
    writer.on('error', reject);
  });
}

const enhanceUserStory = async (story) => {
  const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');
  const headers = {
    "Accept": "application/json",
    "Authorization": `Basic ${auth}`
  };

  const issueDetails = await fetchIssueDetailsWithAttachments(story['key']);
  const attachments = issueDetails.fields.attachment.map(att => { return {content: att.content, filename: att.filename} });
  const fileReferences = [];

  for (const attachment of attachments) {
    // Check if the file reference already exists in the database
    let fileRef = await FileReference.findOne({
      userStoryNumber: story.key,
      filename: attachment.filename
    });

    if (!fileRef) {
      // Fetch the image as a stream using axios if the file reference does not exist
      const localFilePath = await downloadImage(attachment.content, attachment.filename, headers);

      // Upload the stream to OpenAI
      const fileResponse = await openai.files.create({
        file: fs.createReadStream(localFilePath),
        purpose: "vision",
      });

      // Create a new document in MongoDB for the file reference
      fileRef = new FileReference({
          userStoryNumber: story.key,
          openAIFileId: fileResponse.id,
          filename: fileResponse.filename,
          bytes: fileResponse.bytes
      });

      await fileRef.save();
    }

    fileReferences.push(fileRef);
  }

  if (attachments.length > 0) console.log(`${story['key']} has ${attachments.length} attachments added`)

  const prompt = {
    userStoryNumber: story['key'],
    "Epic #": story['fields']['parent'] ? story['fields']['parent']['key'] || "No parent provided." : "No parent provided.",
    "User Story #": story['key'],
    "Title": story['fields']['summary'],
    "Description": extractText(story['fields']['description']) || "No description provided.",
    "Acceptance Criteria": extractText(story['fields']['customfield_10900']) || "No acceptance criteria provided.",
    "Priority": story['fields']['priority']['name'],
    "Developer": story['fields']['assignee']['displayName'],
    "QA": "Ronnel / Brent / Raychal",
    "Product Owner": story['fields']['reporter']['displayName']
  };
  return prompt;
};

const createThreadAndRun = async (data, assistant_id) => {
  // Extract Attachments and prepare the base data without Attachments
  const { Attachments, ...baseData } = data;

  // Initialize the messages array with a single object
  const messages = [
    {
      role: "user",
      content: [
        {
          text: JSON.stringify(baseData),
          type: "text"
        }
      ]
    }
  ];

  // Query FileReference model by userStoryNumber to get attachments
  const fileReferences = await FileReference.find({ userStoryNumber: data.userStoryNumber });

  // If fileReferences is not empty, add each as an image_file type object to the content array
  if (fileReferences && fileReferences.length > 0) {
    fileReferences.forEach(fileRef => {
      messages[0].content.push({
        type: "image_file",
        image_file: {
          file_id: fileRef.openAIFileId // Use openAIFileId from the FileReference model
        }
      });
    });
  }

  // Create and run the thread with the formatted messages
  const threadRun = await openai.beta.threads.createAndRun({
    assistant_id: assistant_id,
    thread: {
      messages: messages
    }
  });

  return threadRun;
};

const createRun = async (thread_id, assistant_id) => {
    const run = await openai.beta.threads.runs.create(thread_id, {assistant_id: assistant_id});
    return run;
}

const retrieveRun = async (run_id, thread_id) => {
    const run = await openai.beta.threads.runs.retrieve(thread_id, run_id);
    return run;
}

const retrieveThreadMessages = async (thread_id) => {
    const thread = await openai.beta.threads.messages.list(thread_id);
    return thread;
}

async function getStoredDataOrGenerate(data, assistantId, model) {
    let generatedData = {}; // Initialize the message
    let messageId = ""; // Initialize the message ID
    let Model = model;
    
    // get the data["User Story #"] and store in variable. if not found get the data.userStoryNumber
    const userStoryNumber = data["User Story #"] || data.userStoryNumber;
    const title = data["Title"] || data.title;
    console.log('Getting stored data or generating new data for ' + userStoryNumber + '...')
  
    // Check if the message is stored
    const storedMessage = await Model.findOne({ userStoryNumber });
    if (storedMessage) {
      console.log('Returning stored message');
      return storedMessage.data;
    }
  
    // If not stored, send to assistant
    const { id, thread_id } = await createThreadAndRun(data, assistantId);

    // Poll the run until it is completed
    let runCompleted = false;
    while (!runCompleted) {
        const statusResponse = await retrieveRun(id, thread_id); // Implement this function based on your API
        if (statusResponse.status === 'completed') {
            const messages = await retrieveThreadMessages(thread_id);
            const lastMessage = messages.data[0];

            generatedData = JSON.parse(lastMessage.content[0].text.value);
            messageId = lastMessage.id;
            runCompleted = true;
        } else {
            // Wait for a bit before polling again
            console.log('...');
            await new Promise(resolve => setTimeout(resolve, 5000)); // Wait for 5 seconds
        }
    }
  
    // Save the new message to the database
    const newEntry = new Model({ userStoryNumber, title, runId: id, threadId: thread_id, messageId, data: generatedData });
    await newEntry.save();
  
    return newEntry.data;
}

async function createIssueLink(inwardIssue, outwardIssue, linkType) {
  const url = jiraURL + "/rest/api/3/issueLink";
  const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');
  const headers = {
    "Accept": "application/json",
    "Content-Type": "application/json",
    "Authorization": `Basic ${auth}`
  };
  const body = {
      // comment: {
      //   body: {
      //     content: [
      //       {
      //         content: [
      //           {
      //             text: "Linked related issue!",
      //             type: "text"
      //           }
      //         ],
      //         type: "paragraph"
      //       }
      //     ],
      //     type: "doc",
      //     version: 1
      //   }
      // },
      inwardIssue: {
        key: inwardIssue
      },
      outwardIssue: {
        key: outwardIssue
      },
      type: {
        name: linkType
      }
    };

  const response = await axios.post(url, body,{ headers });
  console.log(`Created new link for ${inwardIssue} and ${outwardIssue} with linktype ${linkType}`);
  return response;
}

async function createTestIssueIfNotExist(testCaseName, issueType) {
  // Check if test issue exists
  const existingTestCase = await findItemByName(testCaseName, issueType);
  if (existingTestCase.total == 0) {
      const url = jiraURL + "/rest/api/3/issue";
      const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');
      const headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": `Basic ${auth}`
      };
      const body = {
        fields: {
          summary: testCaseName,
          issuetype: {
            id: issueType
          },
          project: {
            key: "WCX"
          }
        }
      };
  
    const response = await axios.post(url, body,{ headers });
    console.log(`Created new test issue: ${response.data.key} with issuetype ${issueType}`);
    return response.data;
  } else {
    console.log(`Test issue already exists: ${testCaseName} with issuetype ${issueType}`);
    return existingTestCase.issues[0];
  }
}

async function createTestIssueInXrayIfNotExist(story, testCase, issueType) {
  try {
    const testCaseName = `[${story['User Story #']}] ${testCase.title}`;
    const unstructured = `*Steps:*\n${testCase.steps.map((step, index) => `${index + 1}. ${step}`).join('\n')}\n\n*Expected Results:*\n${testCase.expectedResults}\n\n*Type:* ${testCase.type}`;

    const existingTestCase = await findItemByName(testCaseName, issueType);
    if (existingTestCase.total == 0) {
      const endpoint = 'https://xray.cloud.getxray.app/api/v2/graphql';

      const mutation = `
        mutation createTest($unstructured: String!, $testCaseName: String!){
          createTest(
            testType: { name: "Generic" },
            unstructured: $unstructured,
            jira: {
              fields: { summary: $testCaseName, project: {key: "WCX"} }
            },
            preconditionIssueIds: ["735970", "735971"]
          ) {
            test {
              issueId
              testType {
                name
              }
              unstructured
              jira(fields: ["key"])
            }
            warnings
          }
        }
      `;

      const variables = {
        unstructured,
        testCaseName
      };

      const response = await axios.post(endpoint, JSON.stringify({
        query: mutation,
        variables
      }), {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${xray_token}`
        }
      });

      if (response.status !== 200) {
        throw new Error(`HTTP error! status: ${response.status} - Body: ${JSON.stringify(response.data)}`);
      }

      const jsonResponse = response.data;

      if (jsonResponse.data) {
        console.log(`Created new test issue: ${JSON.stringify(jsonResponse.data.createTest.test.jira.key)}`);
      } else if (jsonResponse.errors) {
        console.error('Errors:', jsonResponse.errors);
      }

      return jsonResponse;
    } else {
      console.log(`Test issue already exists: ${testCaseName} with issuetype ${issueType}`);
      return existingTestCase.issues[0];
    }
  } catch (error) {
    console.error('Error:', error.message);
  }
}

// Find JIRA Issue by name
async function findItemByName(name, issueTypeId) {
  const url = jiraURL + "/rest/api/3/search";
  const auth = Buffer.from(`${email}:${apiToken}`).toString('base64');
  const headers = {
    "Accept": "application/json",
    "Authorization": `Basic ${auth}`
  };

  // Properly escape special characters in JQL and format the query
  const escapedName = name.replace(/([\\"])/g, '\\$1'); // Escape backslashes and double quotes
  const jql = `project = WCX AND issuetype = ${issueTypeId} AND summary ~ "\\"${escapedName}\\"" ORDER BY created DESC`;

  // Adjust the query parameter to ensure proper formatting
  const response = await axios.get(url, { headers, params: { jql } });
  return response.data;
}

async function addTests(issueId, testIssueIds, actionType) {
  const endpoint = 'https://xray.cloud.getxray.app/api/v2/graphql';
  let mutation;

  switch (actionType) {
    case 'testSet':
      mutation = `
        mutation addTestsToTestSet($issueId: String!, $testIssueIds: [String]!) {
            addTestsToTestSet(issueId: $issueId, testIssueIds: $testIssueIds) {
                warning
                addedTests
            }
        }
      `;
      break;
    case 'testPlan':
      mutation = `
        mutation addTestsToTestPlan($issueId: String!, $testIssueIds: [String]!) {
            addTestsToTestPlan(issueId: $issueId, testIssueIds: $testIssueIds) {
                warning
                addedTests
            }
        }
      `;
      break;
    case 'testExecution':
      mutation = `
        mutation addTestsToTestExecution($issueId: String!, $testIssueIds: [String]!) {
            addTestsToTestExecution(issueId: $issueId, testIssueIds: $testIssueIds) {
                warning
                addedTests
            }
        }
      `;
      break;
    case 'precondition':
      mutation = `
        mutation addTestsToPrecondition($issueId: String!, $testIssueIds: [String]!) {
            addTestsToPrecondition(issueId: $issueId, testIssueIds: $testIssueIds) {
                warning
                addedTests
            }
        }
      `;
    default:
      throw new Error('Invalid action type');
  }

  const variables = {
    issueId,
    testIssueIds
  };

  const response = await axios.post(endpoint, JSON.stringify({
    query: mutation,
    variables
  }), {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${xray_token}`
    }
  });

  if (response.status !== 200) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }

  const jsonResponse = response.data;

  if (jsonResponse.data && jsonResponse.data[`addTestsTo${actionType.charAt(0).toUpperCase() + actionType.slice(1)}`]) {
    console.log(`Tests added successfully to ${actionType}:`, jsonResponse.data[`addTestsTo${actionType.charAt(0).toUpperCase() + actionType.slice(1)}`].addedTests);
  } else if (jsonResponse.errors) {
    console.error('Errors:', jsonResponse.errors);
  }
}

async function addTestExecutions(issueId, testExecIssueIds, actionType) {
  const endpoint = 'https://xray.cloud.getxray.app/api/v2/graphql';
  let mutation;

  switch (actionType) {
    case 'testPlan':
      mutation = `
        mutation addTestExecutionsToTestPlan($issueId: String!, $testExecIssueIds: [String]!) {
            addTestExecutionsToTestPlan(issueId: $issueId, testExecIssueIds: $testExecIssueIds) {
                warning
                addedTestExecutions
            }
        }
      `;
      break;
    case 'test':
      mutation = `
        mutation addTestExecutionsToTests($issueId: String!, $testExecIssueIds: [String]!) {
            addTestExecutionsToTests(issueId: $issueId, testExecIssueIds: $testExecIssueIds) {
                warning
                addedTestExecutions
            }
        }
      `;
      break;
    default:
      throw new Error('Invalid action type');
  }

  const variables = {
    issueId,
    testExecIssueIds
  };

  const response = await axios.post(endpoint, JSON.stringify({
    query: mutation,
    variables
  }), {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${xray_token}`
    }
  });

  if (response.status !== 200) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }

  const jsonResponse = response.data;

  if (jsonResponse.data && jsonResponse.data[`addTestExecsTo${actionType.charAt(0).toUpperCase() + actionType.slice(1)}`]) {
    console.log(`Tests added successfully to ${actionType}:`, jsonResponse.data[`addTestExecsTo${actionType.charAt(0).toUpperCase() + actionType.slice(1)}`].addedTests);
  } else if (jsonResponse.errors) {
    console.error('Errors:', jsonResponse.errors);
  }
}

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const main = async () => {
  let planName = '';
  let planType = '';
  let includeEdgeCases = false;
  let isAutomated = true;

  const testCases = [];

  // Ask for Sprint Number
  planName = await new Promise((resolve) => {
    rl.question('What is the name of the test plan? (ex. Sprint 44) ', (answer) => {
      resolve(answer);
    });
  });

  // Ask for Plan Type (optional)
  planType = await new Promise((resolve) => {
    rl.question('What is the plan type? (Press enter to skip) ', (answer) => {
      resolve(answer);
    });
  });

  // Ask if Edge Cases should be included
  includeEdgeCases = await new Promise((resolve) => {
    rl.question('Include Edge Cases? (y/n) [deafult: n] ', (answer) => {
      resolve(answer.toLowerCase() === 'y');
    });
  });

  // Construct Test Plan Name
  const testPlanName = planType ? `${planName} - ${planType}` : planName;

  // GENERATE TEST PLAN ISSUE IN JIRA
  const testPlanIssue = await createTestIssueIfNotExist(testPlanName, issueTypes.TEST_PLAN);

  // GET JIRA USER STORIES
  const userStories = await fetchJiraUserStories();
  // console.log(JSON.stringify(userStories, null, 2));

  // PROCESS USER STORIES
  // Initialize an empty array to hold the formatted stories
  const formattedStories = [];

  // Loop through each story in userStories['issues']
  for (const story of userStories['issues']) {
    // Await the result of enhanceUserStory and push it into the formattedStories array
    const formattedStory = await enhanceUserStory(story);
    formattedStories.push(formattedStory);
  }

  // GENERATE TEST ISSUES FOR EACH USER STORY
  for (const story of formattedStories) {
    // Search MongoDB for an existing document with the userStoryNumber
    const existingDocument = await QATestGenerator.findOne({ userStoryNumber: story['User Story #'] });
    let testIssues = existingDocument ? existingDocument.testIssues : [];

    if (testIssues.length === 0) {
      // Create a Test Set for each User Story
      const testExecutionName = `${planName} - [${story['User Story #']}] - Automated`;
      const testExecutionIssue = await createTestIssueIfNotExist(testExecutionName, issueTypes.TEST_EXECUTION);
      const testSetIssue = await createTestIssueIfNotExist(`${planName} - [${story['User Story #']}] - Test Set`, issueTypes.TEST_SET);

      // LINK TEST PLAN AND TEST EXECUTION
      await addTestExecutions(testPlanIssue.id, [testExecutionIssue.id], "testPlan");

      // Link User Story to Test Plan, Test Execution, and Test Set
      const storyToTestPlanLinkResponse = await createIssueLink(testPlanIssue.key, story['User Story #'], "Test");
      const storyToTestExecutionLinkResponse = await createIssueLink(testExecutionIssue.key, story['User Story #'], "Test");
      const storyToTestSetLinkResponse = await createIssueLink(testSetIssue.key, story['User Story #'], "Test");

      if (storyToTestPlanLinkResponse.status == 201)  {
        console.log(`Successfully linked ${story['User Story #']} to ${testPlanIssue.key}`);
      } else {
        console.error(`Failed to link ${story['User Story #']} to ${testPlanIssue.key}`);
      }

      if (storyToTestExecutionLinkResponse.status == 201)  {
        console.log(`Successfully linked ${story['User Story #']} to ${testExecutionIssue.key}`);
      } else {
        console.error(`Failed to link ${story['User Story #']} to ${testExecutionIssue.key}`);
      }

      if (storyToTestSetLinkResponse.status == 201)  {
        console.log(`Successfully linked ${story['User Story #']} to ${testSetIssue.key}`);
      } else {
        console.error(`Failed to link ${story['User Story #']} to ${testSetIssue.key}`);
      }

      // Generate test cases for each User Story
      const savedJiraTicketHelperData = await getStoredDataOrGenerate(story, jiraTicketHelperAssistantId, JiraTicketHelper);
      const savedQATestGeneratorData = await getStoredDataOrGenerate(savedJiraTicketHelperData, qaTestCaseGeneratorAssistantId, QATestGenerator);
      
      // Create test issues for each test case
      for (const testCase of savedQATestGeneratorData.testCases) {
        const testIssue = await createTestIssueInXrayIfNotExist(story, testCase, issueTypes.TEST);
        testIssues.push(testIssue.data.createTest.test.issueId);
      }

      if (includeEdgeCases) {
        // Create test issues for each edge case
        for (const edgeCase of savedQATestGeneratorData.edgeCases) {
          const testIssue = await createTestIssueInXrayIfNotExist(story, edgeCase, issueTypes.TEST);
          testIssues.push(testIssue.data.createTest.test.issueId);
        }
        console.log("Included edge cases in test issues.");
      }

      // Link test issues to the Test Set and Test Execution
      await addTests(testSetIssue.id, testIssues, "testSet");
      await addTests(testExecutionIssue.id, testIssues, "testExecution");
      await addTests(testPlanIssue.id, testIssues, "testPlan");
      // Add to default preconditions
      await addTests('735970', testIssues, "precondition");
      await addTests('735971', testIssues, "precondition");
      
      savedQATestGeneratorData.title = story['Title'];
      testCases.push(savedQATestGeneratorData);

      // Save the updated data back to MongoDB
      await QATestGenerator.updateOne(
        { userStoryNumber: story['User Story #'] }, // Assuming userStoryNumber is the unique identifier
        { $set: { testIssues: testIssues } }
      );

    } else {
      console.log(`Test issues already exist for ${story['User Story #']}`);
      const data = existingDocument.data;
      data['title'] = existingDocument.title;
      testCases.push(data);
      continue;
    }
  }

  const excelGenerator = new TestCaseDocGenerator(testCases);
  excelGenerator.generateTestCaseDoc(`${testPlanName}_test_case_document.xlsx`)

  console.log('All user stories have been processed.');

  // close the mongoose connection
  mongoose.connection.close();
};

main().catch(console.error);