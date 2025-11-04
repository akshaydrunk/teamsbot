# Teams Notification Bot

This is a Python-based notification bot for Microsoft Teams. It can be installed in Teams channels, group chats, and personal chats to send proactive notifications.

## How it Works

When the bot is added to a Team, channel, or chat, Microsoft Teams sends a notification to the bot. The bot then stores the necessary conversation reference information in a file named `recipients.json`. This file is automatically created or updated whenever the bot is added to or removed from a new conversation.

This `recipients.json` file is then used by the bot to send proactive messages to the appropriate channels or users.

## Installation and Setup

### 1. Configuration

To run the bot, you need to configure it with a `BOT_ID` and `APP_PASSWORD` from the Microsoft Teams Developer Portal.

#### 1.1. Getting `BOT_ID` and `APP_PASSWORD`

1.  **Go to the Teams Developer Portal:** [https://dev.teams.microsoft.com/](https://dev.teams.microsoft.com/)
2.  **Select your App:** Navigate to the app you have created.
3.  **Get the `BOT_ID`:** The `BOT_ID` is your **App ID**. You can find this in the "Basic information" section or in your `manifest.json` file.
4.  **Get the `APP_PASSWORD`:**
    *   Go to the "App passwords" section for your bot.
    *   Click on "Add an app password" to generate a new password.
    *   **Important:** Copy this password and save it somewhere safe. You will not be able to see it again.

#### 1.2. Updating `teamsbot.py`

Open the `teamsbot.py` file and update the following lines with the values you obtained:

```python
# Configuration - Replace with your actual Bot ID and App Password
BOT_ID = "YOUR_BOT_ID_HERE"  # From your manifest
APP_PASSWORD = "YOUR_APP_PASSWORD_HERE"  # From Teams Developer Portal
```

### 2. Exposing the Bot to the Internet (using ngrok)

For the initial setup, the bot needs to be accessible from the public internet so that Teams can send installation events. A simple way to do this is by using `ngrok`.

1.  **Start the bot:**
    ```bash
    python teamsbot.py
    ```
2.  **Expose the bot's port (3978) using ngrok:**
    ```bash
    ngrok http 3978
    ```
3.  **Copy the ngrok URL:** `ngrok` will provide a public HTTPS URL (e.g., `https://xxxx-xx-xx-xx-xx.ngrok-free.app`).
4.  **Update your bot's messaging endpoint:** In the Teams Developer Portal, go to your app, and under "App features", update the messaging endpoint to `YOUR_NGROK_URL/api/messages`.

### 3. Uploading the Bot to Teams

1.  **Download the App Package:** In the Teams Developer Portal, go to your app and select "Distribute". Download the app package (it will be a zip file).
2.  **Upload the App to Teams:**
    *   In Microsoft Teams, go to "Apps".
    *   Click on "Manage your apps" -> "Upload an app" -> "Upload a custom app".
    *   Select the zip file you downloaded.

### 4. Adding the Bot and Generating Recipients

Now, add the bot to a Team, channel, or chat. When the bot is successfully added, it will receive an event from Teams, and the `recipients.json` file will be created or updated with the new conversation details.

### 5. Switching to Internal Network

Once the `recipients.json` file has been populated, you no longer need `ngrok` to send notifications. You can stop `ngrok` and send notifications by calling the bot's API on your internal network (e.g., `http://localhost:3978/send`).

## Available APIs

The bot exposes the following API endpoints:

*   `POST /api/messages`: The main endpoint for receiving activities from Teams.
*   `POST /send`: Sends a notification to specified recipients.
*   `GET /status`: Retrieves the bot's status and a list of recipients.
*   `GET /targets`: Lists all available targeting options.
*   `GET /health`: A health check endpoint.

## API Usage Examples

You can use `curl` or any other API client to interact with the bot's APIs.

### Send a Notification

This endpoint sends a notification to all or a targeted set of recipients.

**Send a simple message to all recipients:**

```bash
curl -X POST http://localhost:3978/send \
-H "Content-Type: application/json" \
-d 
'{ "message": "This is a test notification to everyone." }'
```

**Send a message to specific conversation IDs:**

```bash
curl -X POST http://localhost:3978/send \
-H "Content-Type: application/json" \
-d 
'{ "message": "This is a targeted notification.", "conversation_ids": ["19:xxxxxxxxxxxxxxxxxxxx@thread.tacv2"] }'
```

**Send a message to recipients with specific tags:**

```bash
curl -X POST http://localhost:3978/send \
-H "Content-Type: application/json" \
-d 
'{ "message": "This is a notification for the \'channel\' tag.", "tags": ["channel"] }'
```

### Get Bot Status

This endpoint retrieves the bot's status and a list of all recipients.

```bash
curl http://localhost:3978/status
```

### List Targeting Options

This endpoint lists all available targeting options, including conversation IDs, tags, teams, and channels.

```bash
curl http://localhost:3978/targets
```
