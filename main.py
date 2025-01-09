import os
import asyncio
from datetime import datetime, timedelta
import pandas as pd
import discord
from discord.ext import tasks
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')

# Create the bot client
intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
client = discord.Client(intents=intents)

# Global variable to hold the schedule data
schedule_data = None


@client.event
async def on_ready():
    print(f"{client.user} has connected to Discord!")
    check_schedule.start()  # Start the task that checks for reminders


@client.event
async def on_message(message):
    global schedule_data

    # Ignore messages from the bot itself
    if message.author == client.user:
        return

    # Check for Excel file attachments
    if message.attachments:
        for attachment in message.attachments:
            if attachment.filename.endswith(('.xls', '.xlsx')):
                try:
                    # Download the file
                    file_path = f"./{attachment.filename}"
                    await attachment.save(file_path)

                    # Load the Excel file into a DataFrame
                    schedule_data = pd.read_excel(file_path)

                    # Send confirmation and preview
                    preview = schedule_data.head().to_string(index=False)
                    await message.channel.send(
                        f"Schedule loaded successfully! Here's a preview:\n```{preview}```"
                    )

                except Exception as e:
                    await message.channel.send(f"Error reading the Excel file: {str(e)}")
            else:
                await message.channel.send("I require an Excel file and can't read any other one!")


@tasks.loop(seconds=30)
async def check_schedule():
    """
    Periodically checks the schedule and sends reminders if an event is upcoming.
    """
    global schedule_data

    if schedule_data is not None:
        try:
            now = datetime.now()
            current_day = now.strftime("%a")  # Mon, Tue, etc.
            current_time = now.strftime("%I:%M %p")  # HH:MM AM/PM

            # Iterate through the schedule to check for reminders
            for item in schedule_data.iloc():
                event_day, event_time = item['Days/Times'].split()
                course_name = item['Course Name']

                # Check if it's 10 minutes before the event
                event_datetime = datetime.strptime(event_time, "%I:%M %p") - timedelta(minutes=10)
                reminder_time = event_datetime.strftime("%I:%M %p")

                if event_day == current_day and reminder_time == current_time:
                    # Send reminder
                    # Note: Replace `user_id` with a valid ID if DMs are needed
                    channel = discord.utils.get(client.get_all_channels(), name="general")
                    if channel:
                        await channel.send(
                            f"Reminder: Your `{course_name}` class starts in 10 minutes!"
                        )

        except Exception as e:
            print(f"Error in check_schedule: {str(e)}")


# Run the bot
async def main():
    await client.start(TOKEN)


if __name__ == "__main__":
    asyncio.run(main())