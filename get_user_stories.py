import requests
from requests.auth import HTTPBasicAuth
import json
import os
from dotenv import load_dotenv  # Import load_dotenv from python-dotenv
import time

from openai import OpenAI

# Load environment variables from .env file
load_dotenv()

# import pandas as pd

# Initialize OpenAI
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)
jira_ticket_helper_assistant_id = "asst_4Vl3xpvP7T1BRFzypol7rs8Y"
qa_test_case_generator_assistant_id = "asst_AHfgXfV1JN1v6qU57h5I6dXc"

# Retrieve environment variables
email = os.getenv("JIRA_EMAIL")
api_token = os.getenv("JIRA_API_TOKEN")

def fetch_jira_user_stories():
    url = "https://concentrixdev.atlassian.net/rest/api/3/search"
    auth = HTTPBasicAuth(email, api_token)
    headers = {
      "Accept": "application/json"
    }

    query = {
      'jql': 'project = WCX AND issuetype = Story AND sprint in openSprints() AND status NOT IN (Closed, Cancelled) ORDER BY created DESC'
    }

    response = requests.request(
       "GET",
       url,
       headers=headers,
       params=query,
       auth=auth
    )
    
    return response.json()


def extract_text_recursive(content_item):
    text_parts = []
    if isinstance(content_item, dict):
        # If the item is a dictionary, check for 'text' and 'content' keys
        if "text" in content_item:
            text_parts.append(content_item["text"])
        elif "content" in content_item:
            for nested_content in content_item["content"]:
                text_parts += extract_text_recursive(nested_content)
    elif isinstance(content_item, list):
        # If the item is a list, iterate through its elements
        for item in content_item:
            text_parts += extract_text_recursive(item)
    return text_parts


def extract_text(description):
    text_parts = []

    if "content" in description:
        for content_item in description["content"]:
            text_parts += extract_text_recursive(content_item)

    # Join all parts of the text into a single string with newline characters
    full_text = "\n".join(text_parts)
    return full_text


def enhance_user_story(story):
    prompt = {
        "Epic #": story['fields']['parent']['key'],  # Adjust the custom field ID
        "User Story #": story['key'],
        "Title": story['fields']['summary'],
        "Description": extract_text(story['fields']['description']) or "No description provided.",
        "Acceptance Criteria": extract_text(story['fields'].get('customfield_10900', "No acceptance criteria provided.")),  # Adjust the custom field ID
        "Priority": story['fields']['priority']['name'],
        "Developer": story['fields']['assignee']['displayName'],
        "QA": "Ronnel / Brent / Raychal",  # Adjust as necessary
        "Product Owner": story['fields']['reporter']['displayName'],
    }

    # response = openai.Completion.create(
    #     engine="text-davinci-002",
    #     prompt=json.dumps(prompt),
    #     max_tokens=150
    # )
    
    # return json.loads(response.choices[0].text)
    return prompt


def jira_ticket_helper_create_thread_and_run(user_story):
    user_story_thread_run = client.beta.threads.create_and_run(
        assistant_id=jira_ticket_helper_assistant_id,
        thread={
            "messages": [
                {"role": "user", "content": json.dumps(user_story)},
            ]
        })
    return user_story_thread_run

user_stories = fetch_jira_user_stories()
# print(json.dumps(user_stories, sort_keys=True, indent=4, separators=(",", ": ")))

formatted_stories = [enhance_user_story(story) for story in user_stories['issues']]

user_story_thread_run = jira_ticket_helper_create_thread_and_run(formatted_stories[0])

print(user_story_thread_run)
