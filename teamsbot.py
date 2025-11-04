import os
import json
import asyncio
from datetime import datetime
from typing import Dict, List, Any, Optional

from aiohttp import web
from aiohttp.web import Request, Response, json_response
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    TurnContext,
    ActivityHandler,
    MessageFactory
)
from botbuilder.schema import (
    Activity,
    ActivityTypes,
    ChannelAccount,
    ConversationAccount,
    ConversationParameters,
    ConversationReference
)

# Configuration - Replace with your actual Bot ID and App Password
BOT_ID = ""  # From your manifest
APP_PASSWORD = ""  # From Teams Developer Portal
RECIPIENTS_FILE = "recipients.json"

class NotificationBot(ActivityHandler):
    """
    Notification-only bot that stores recipient information and sends proactive messages
    """

    def __init__(self):
        super().__init__()
        self.recipients = self._load_recipients()
        self._processed_installations = set()  # Track processed installations to avoid duplicates

    def _load_recipients(self) -> Dict[str, Any]:
        """Load recipients from file"""
        try:
            if os.path.exists(RECIPIENTS_FILE):
                with open(RECIPIENTS_FILE, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading recipients: {e}")
        return {}

    def _save_recipients(self):
        """Save recipients to file"""
        try:
            with open(RECIPIENTS_FILE, 'w') as f:
                json.dump(self.recipients, f, indent=2)
            print(f"Saved {len(self.recipients)} recipients to {RECIPIENTS_FILE}")
        except Exception as e:
            print(f"Error saving recipients: {e}")

    async def on_turn(self, turn_context: TurnContext):
        """Handle all incoming activities"""
        print(f"=== INCOMING ACTIVITY ===")
        print(f"Activity type: {turn_context.activity.type}")
        print(f"Activity: {turn_context.activity}")

        # Call the parent handler
        await super().on_turn(turn_context)

    async def on_installation_update_add(self, turn_context: TurnContext):
        """Handle Teams installation update event - PRIMARY installation handler"""
        try:
            print(f"=== INSTALLATION UPDATE ADD ===")
            print(f"Activity: {turn_context.activity}")
            
            conversation_id = turn_context.activity.conversation.id
            
            # Check if we've already processed this installation
            if conversation_id in self._processed_installations:
                print(f"Installation already processed for {conversation_id}, skipping...")
                return
            
            # Mark as processed
            self._processed_installations.add(conversation_id)
            
            # Store recipient and send welcome message
            await self._store_recipient(turn_context)
            
        except Exception as e:
            print(f"Error handling installation update: {e}")

    async def on_members_added_activity(self, members_added: List[ChannelAccount], turn_context: TurnContext):
        """Handle bot installation - SECONDARY handler (disabled to prevent duplicates)"""
        try:
            print(f"=== MEMBERS ADDED EVENT ===")
            print(f"Members added: {members_added}")
            
            # Check if the bot was added
            bot_member = next((member for member in members_added
                             if member.id == BOT_ID or member.id == f"28:{BOT_ID}"), None)

            if bot_member:
                conversation_id = turn_context.activity.conversation.id
                
                # Only process if not already handled by installation_update_add
                if conversation_id not in self._processed_installations:
                    print(f"Processing installation via members_added for {conversation_id}")
                    self._processed_installations.add(conversation_id)
                    await self._store_recipient(turn_context)
                else:
                    print(f"Installation already processed via installation_update_add, skipping...")

        except Exception as e:
            print(f"Error handling member added: {e}")

    async def _store_recipient(self, turn_context: TurnContext):
        """Store recipient information for proactive messaging"""
        try:
            # Get conversation reference for proactive messaging
            conversation_ref = TurnContext.get_conversation_reference(turn_context.activity)

            # Extract additional Teams-specific information
            conversation = turn_context.activity.conversation
            channel_data = getattr(turn_context.activity, 'channel_data', {}) or {}
            team_info = channel_data.get('team', {}) if isinstance(channel_data, dict) else {}
            channel_info = channel_data.get('channel', {}) if isinstance(channel_data, dict) else {}

            # Extract recipient information with enhanced metadata
            recipient_info = {
                "conversation_id": conversation.id,
                "conversation_type": conversation.conversation_type,
                "conversation_name": conversation.name,
                "service_url": turn_context.activity.service_url,
                "channel_id": turn_context.activity.channel_id,
                "tenant_id": getattr(conversation, 'tenant_id', None),

                # Teams-specific information
                "team_id": team_info.get('id') if team_info else None,
                "team_name": team_info.get('name') if team_info else None,
                "channel_name": channel_info.get('name') if channel_info else None,
                "teams_channel_id": channel_info.get('id') if channel_info else None,

                # For easy identification
                "display_name": self._generate_display_name(conversation, team_info, channel_info),
                "tags": self._generate_tags(conversation, team_info, channel_info),

                "conversation_reference": {
                    "activity_id": conversation_ref.activity_id,
                    "bot": {
                        "id": conversation_ref.bot.id,
                        "name": conversation_ref.bot.name
                    } if conversation_ref.bot else None,
                    "channel_id": conversation_ref.channel_id,
                    "conversation": {
                        "conversation_type": conversation_ref.conversation.conversation_type,
                        "id": conversation_ref.conversation.id,
                        "is_group": conversation_ref.conversation.is_group,
                        "name": conversation_ref.conversation.name,
                        "tenant_id": getattr(conversation_ref.conversation, 'tenant_id', None)
                    } if conversation_ref.conversation else None,
                    "service_url": conversation_ref.service_url,
                    "user": {
                        "id": conversation_ref.user.id,
                        "name": conversation_ref.user.name
                    } if conversation_ref.user else None
                },
                "added_at": datetime.utcnow().isoformat()
            }
            print(f"Recipient Info: {recipient_info}")

            # Store recipient with conversation ID as key
            self.recipients[conversation.id] = recipient_info
            self._save_recipients()

            print(f"Bot installed in: {recipient_info['display_name']}")
            print(f"Conversation ID: {conversation.id}")
            print(f"Tags: {recipient_info['tags']}")

            # Send single welcome message
            welcome_message = MessageFactory.text("Notify bot added successfully!")
            await turn_context.send_activity(welcome_message)

        except Exception as e:
            print(f"Error storing recipient: {e}")

    def _generate_display_name(self, conversation, team_info, channel_info):
        """Generate a human-readable display name for the conversation"""
        if conversation.conversation_type == "channel" and team_info and channel_info:
            return f"{team_info.get('name', 'Unknown Team')} > {channel_info.get('name', 'Unknown Channel')}"
        elif conversation.conversation_type == "personal":
            return f"Personal Chat ({conversation.name or 'Unknown User'})"
        elif conversation.conversation_type == "groupChat":
            return f"Group Chat ({conversation.name or 'Unnamed Group'})"
        else:
            return f"{conversation.conversation_type.title()} ({conversation.name or conversation.id[:20]}...)"

    def _generate_tags(self, conversation, team_info, channel_info):
        """Generate searchable tags for the conversation"""
        tags = [conversation.conversation_type]

        if team_info:
            tags.append(f"team:{team_info.get('name', '').lower().replace(' ', '-')}")

        if channel_info:
            tags.append(f"channel:{channel_info.get('name', '').lower().replace(' ', '-')}")

        if conversation.name:
            tags.append(f"name:{conversation.name.lower().replace(' ', '-')}")

        return [tag for tag in tags if tag and ':' in tag or tag in ['channel', 'personal', 'groupChat']]

    async def on_members_removed_activity(self, members_removed: List[ChannelAccount], turn_context: TurnContext):
        """Handle bot removal - clean up recipient information"""
        try:
            # Check if the bot was removed (Teams prefixes bot ID with "28:")
            bot_member = next((member for member in members_removed
                             if member.id == BOT_ID or member.id == f"28:{BOT_ID}"), None)

            if bot_member:
                conversation_id = turn_context.activity.conversation.id
                if conversation_id in self.recipients:
                    display_name = self.recipients[conversation_id].get('display_name', conversation_id)
                    del self.recipients[conversation_id]
                    self._save_recipients()
                    # Remove from processed set
                    self._processed_installations.discard(conversation_id)
                    print(f"Bot removed from: {display_name}")

        except Exception as e:
            print(f"Error handling member removed: {e}")

    async def on_conversation_update_activity(self, turn_context: TurnContext):
        """Handle conversation update activities"""
        print(f"=== CONVERSATION UPDATE ===")
        print(f"Activity: {turn_context.activity}")
        print(f"Members added: {turn_context.activity.members_added}")
        print(f"Members removed: {turn_context.activity.members_removed}")

        # Call the parent handler to trigger members_added/removed events
        await super().on_conversation_update_activity(turn_context)

    async def on_message_activity(self, turn_context: TurnContext):
        """Handle incoming messages (notification-only bot, so we don't process these)"""
        print(f"=== MESSAGE ACTIVITY ===")
        print(f"Received message: {turn_context.activity.text}")
        print(f"From: {turn_context.activity.from_property}")
        print(f"Conversation: {turn_context.activity.conversation.id}")
        pass

class TeamsNotificationServer:
    """
    HTTP server that hosts the bot and provides API endpoints
    """

    def __init__(self):
        # Bot Framework Adapter
        self.settings = BotFrameworkAdapterSettings(BOT_ID, APP_PASSWORD)
        self.adapter = BotFrameworkAdapter(self.settings)

        # Bot instance
        self.bot = NotificationBot()

        # Error handler
        async def on_error(context: TurnContext, error: Exception):
            print(f"Error: {error}")
            await context.send_activity(MessageFactory.text(f"Sorry, an error occurred: {str(error)}"))

        self.adapter.on_turn_error = on_error

    async def messages_handler(self, request: Request) -> Response:
        """Handle incoming messages from Teams"""
        try:
            # Get the request body
            body = await request.text()

            # Parse activity
            activity = Activity().deserialize(json.loads(body))

            # Create auth header
            auth_header = request.headers.get("Authorization", "")

            # Process the activity
            await self.adapter.process_activity(activity, auth_header, self.bot.on_turn)

            return Response(status=200)

        except Exception as e:
            print(f"Error processing message: {e}")
            return Response(status=500, text=str(e))

    async def send_notification_handler(self, request: Request) -> Response:
        """Send proactive notification to all recipients or targeted recipients"""
        try:
            # Get request data
            data = await request.json()
            message_text = data.get('message', 'Default notification message')

            # Targeting options
            target_conversation_ids = data.get('conversation_ids', [])  # List of specific conversation IDs
            target_tags = data.get('tags', [])  # List of tags to match
            target_teams = data.get('teams', [])  # List of team names
            target_channels = data.get('channels', [])  # List of channel names
            exclude_conversation_ids = data.get('exclude_conversation_ids', [])  # Exclude specific conversations

            # Load current recipients
            recipients = self.bot._load_recipients()

            if not recipients:
                return json_response({"error": "No recipients found"}, status=400)

            # Filter recipients based on targeting criteria
            filtered_recipients = self._filter_recipients(
                recipients,
                target_conversation_ids,
                target_tags,
                target_teams,
                target_channels,
                exclude_conversation_ids
            )

            if not filtered_recipients:
                return json_response({
                    "error": "No recipients match the targeting criteria",
                    "total_recipients": len(recipients),
                    "criteria": {
                        "conversation_ids": target_conversation_ids,
                        "tags": target_tags,
                        "teams": target_teams,
                        "channels": target_channels,
                        "exclude_conversation_ids": exclude_conversation_ids
                    }
                }, status=400)

            sent_count = 0
            errors = []
            sent_to = []

            # Send message to each filtered recipient
            for conversation_id, recipient_info in filtered_recipients.items():
                try:
                    # Reconstruct conversation reference
                    conv_ref_data = recipient_info.get('conversation_reference', {})

                    # Create conversation reference
                    conversation_ref = ConversationReference(
                        activity_id=conv_ref_data.get('activity_id'),
                        bot=ChannelAccount(
                            id=conv_ref_data.get('bot', {}).get('id'),
                            name=conv_ref_data.get('bot', {}).get('name')
                        ) if conv_ref_data.get('bot') else None,
                        channel_id=conv_ref_data.get('channel_id'),
                        conversation=ConversationAccount(
                            conversation_type=conv_ref_data.get('conversation', {}).get('conversation_type'),
                            id=conv_ref_data.get('conversation', {}).get('id'),
                            is_group=conv_ref_data.get('conversation', {}).get('is_group'),
                            name=conv_ref_data.get('conversation', {}).get('name'),
                            tenant_id=conv_ref_data.get('conversation', {}).get('tenant_id')
                        ) if conv_ref_data.get('conversation') else None,
                        service_url=conv_ref_data.get('service_url'),
                        user=ChannelAccount(
                            id=conv_ref_data.get('user', {}).get('id'),
                            name=conv_ref_data.get('user', {}).get('name')
                        ) if conv_ref_data.get('user') else None
                    )

                    # Send proactive message
                    await self.adapter.continue_conversation(
                        conversation_ref,
                        lambda turn_context: self._send_proactive_message(turn_context, message_text),
                        BOT_ID
                    )

                    sent_count += 1
                    sent_to.append({
                        "conversation_id": conversation_id,
                        "display_name": recipient_info.get('display_name'),
                        "tags": recipient_info.get('tags', [])
                    })
                    print(f"Sent notification to: {recipient_info.get('display_name', conversation_id)}")

                except Exception as e:
                    error_msg = f"Failed to send to {recipient_info.get('display_name', conversation_id)}: {str(e)}"
                    errors.append(error_msg)
                    print(error_msg)

            result = {
                "sent_count": sent_count,
                "total_recipients": len(recipients),
                "filtered_recipients": len(filtered_recipients),
                "sent_to": sent_to,
                "errors": errors,
                "targeting_criteria": {
                    "conversation_ids": target_conversation_ids,
                    "tags": target_tags,
                    "teams": target_teams,
                    "channels": target_channels,
                    "exclude_conversation_ids": exclude_conversation_ids
                }
            }

            return json_response(result)

        except Exception as e:
            print(f"Error sending notifications: {e}")
            return json_response({"error": str(e)}, status=500)

    def _filter_recipients(self, recipients, target_conversation_ids, target_tags, target_teams, target_channels, exclude_conversation_ids):
        """Filter recipients based on targeting criteria"""
        filtered = {}

        for conversation_id, recipient_info in recipients.items():
            # Skip if explicitly excluded
            if conversation_id in exclude_conversation_ids:
                continue

            # If no targeting criteria specified, include all (except excluded)
            if not any([target_conversation_ids, target_tags, target_teams, target_channels]):
                filtered[conversation_id] = recipient_info
                continue

            # Check specific conversation IDs
            if target_conversation_ids and conversation_id in target_conversation_ids:
                filtered[conversation_id] = recipient_info
                continue

            # Check tags
            if target_tags:
                recipient_tags = recipient_info.get('tags', [])
                if any(tag in recipient_tags for tag in target_tags):
                    filtered[conversation_id] = recipient_info
                    continue

            # Check team names
            if target_teams:
                team_name = recipient_info.get('team_name', '').lower()
                if any(team.lower() in team_name for team in target_teams):
                    filtered[conversation_id] = recipient_info
                    continue

            # Check channel names
            if target_channels:
                channel_name = recipient_info.get('channel_name', '').lower()
                if any(channel.lower() in channel_name for channel in target_channels):
                    filtered[conversation_id] = recipient_info
                    continue

        return filtered

    async def _send_proactive_message(self, turn_context: TurnContext, message_text: str):
        """Send the actual proactive message"""
        message = MessageFactory.text(f"{message_text}")
        await turn_context.send_activity(message)

    async def status_handler(self, request: Request) -> Response:
        """Get bot status and recipient count with detailed information"""
        recipients = self.bot._load_recipients()

        status = {
            "bot_id": BOT_ID,
            "recipients_count": len(recipients),
            "recipients": [
                {
                    "conversation_id": conv_id,
                    "display_name": info.get('display_name'),
                    "conversation_type": info.get('conversation_type'),
                    "team_name": info.get('team_name'),
                    "channel_name": info.get('channel_name'),
                    "tags": info.get('tags', []),
                    "added_at": info.get('added_at')
                }
                for conv_id, info in recipients.items()
            ]
        }

        return json_response(status)

    async def list_targets_handler(self, request: Request) -> Response:
        """List all available targeting options"""
        recipients = self.bot._load_recipients()

        # Collect all unique targeting options
        conversation_ids = list(recipients.keys())
        all_tags = set()
        teams = set()
        channels = set()

        for info in recipients.values():
            all_tags.update(info.get('tags', []))
            if info.get('team_name'):
                teams.add(info['team_name'])
            if info.get('channel_name'):
                channels.add(info['channel_name'])

        targeting_options = {
            "conversation_ids": conversation_ids,
            "available_tags": sorted(list(all_tags)),
            "available_teams": sorted(list(teams)),
            "available_channels": sorted(list(channels)),
            "recipients_summary": [
                {
                    "conversation_id": conv_id,
                    "display_name": info.get('display_name'),
                    "tags": info.get('tags', [])
                }
                for conv_id, info in recipients.items()
            ]
        }

        return json_response(targeting_options)


def create_app():
    """Create the aiohttp web application"""
    server = TeamsNotificationServer()

    app = web.Application()

    # Teams webhook endpoint
    app.router.add_post('/api/messages', server.messages_handler)

    # API endpoints
    app.router.add_post('/send', server.send_notification_handler)
    app.router.add_get('/status', server.status_handler)
    app.router.add_get('/targets', server.list_targets_handler)


    # Health check
    app.router.add_get('/health', lambda request: json_response({"status": "healthy"}))

    return app

if __name__ == "__main__":
    # Validate configuration
    if APP_PASSWORD == "YOUR_APP_PASSWORD_HERE":
        print("Please set your APP_PASSWORD in the code!")
        print("Get this from the Teams Developer Portal -> Your App -> App passwords")
        exit(1)

    print("Starting Teams Notification Bot with Targeting Support...")
    print(f"Bot ID: {BOT_ID}")
    print(f"Recipients file: {RECIPIENTS_FILE}")
    print("\nEndpoints:")
    print("  POST /api/messages - Teams webhook")
    print("  POST /send - Send notification (with targeting)")
    print("  GET /status - Bot status with recipients")
    print("  GET /targets - List targeting options")

    print("  GET /health - Health check")

    # Create and run the app
    app = create_app()
    web.run_app(app, host='0.0.0.0', port=3978)
