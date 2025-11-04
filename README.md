This is a Python-based notification bot for Microsoft Teams. It can be installed in Teams channels, group chats, and personal chats to send proactive notifications.

## Configuration

To run the bot, you need to configure it with a `BOT_ID` and `APP_PASSWORD` from the Microsoft Teams Developer Portal.

### 1. Getting `BOT_ID` and `APP_PASSWORD`

1.  **Go to the Teams Developer Portal:** [https://dev.teams.microsoft.com/](https://dev.teams.microsoft.com/)
2.  **Select your App:** Navigate to the app you have created.
3.  **Get the `BOT_ID`:** The `BOT_ID` is your **App ID**. You can find this in the "Basic information" section or in your `manifest.json` file.
4.  **Get the `APP_PASSWORD`:**
    *   Go to the "App passwords" section for your bot.
    *   Click on "Add an app password" to generate a new password.
    *   **Important:** Copy this password and save it somewhere safe. You will not be able to see it again.

### 2. Updating `teamsbot.py`

Open the `teamsbot.py` file and update the following lines with the values you obtained:

```python
# Configuration - Replace with your actual Bot ID and App Password
BOT_ID = "YOUR_BOT_ID_HERE"  # From your manifest
APP_PASSWORD = "YOUR_APP_PASSWORD_HERE"  # From Teams Developer Portal
```

## Running the Bot

Once you have configured the `BOT_ID` and `APP_PASSWORD`, you can run the bot:

```bash
python teamsbot.py
```

The bot will start and listen for incoming requests on port 3978.

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
-d '{
    "message": "This is a test notification to everyone."
}'
```

**Send a message to specific conversation IDs:**

```bash
curl -X POST http://localhost:3978/send \
-H "Content-Type: application/json" \
-d '{
    "message": "This is a targeted notification.",
    "conversation_ids": ["19:xxxxxxxxxxxxxxxxxxxx@thread.tacv2"]
}'
```

**Send a message to recipients with specific tags:**

```bash
curl -X POST http://localhost:3978/send \
-H "Content-Type: application/json" \
-d '{
    "message": "This is a notification for the \'channel\' tag.",
    "tags": ["channel"]
}'
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
