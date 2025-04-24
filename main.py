import discord
from discord.ext import commands
import pytesseract
from PIL import Image
import json
import io
import re
import os
import pandas as pd

intents = discord.Intents.default()
intents.messages = True
intents.message_content = True
bot = commands.Bot(command_prefix="?", intents=intents)

def load_existing_data(): # Loads existing data
    try:
        if os.path.exists('data.json'):
            with open('data.json', 'r', encoding='utf-8') as f:
                return json.load(f)
        return []
    except json.JSONDecodeError:
        return []
    
def save_data(new_entries): # Saves data
    existing_data = load_existing_data()

    existing_names = {entry['name'] for entry in existing_data}

    for new_entry in new_entries:
        if new_entry['name'] in existing_names:
            for i, entry in enumerate(existing_data):
                if entry['name'] == new_entry['name']:
                    existing_data[i] = new_entry
                    break
        else:
            existing_data.append(new_entry)

    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(existing_data, f, indent=4, ensure_ascii=False)

    return existing_data

@bot.event
async def on_ready(): # Notifies about connection to discord
    print(f'{bot.user} has connected to Discord!')

@bot.command(name='AMcheck') # Takes picture and reads Name, Score and Amount of tasks done
async def process_leaderboard(ctx):
    if not ctx.message.attachments:
        await ctx.send("Please attach an image with the leaderboard.")
        return

    try:
        attachment = ctx.message.attachments[0]
        image_bytes = await attachment.read()
        image = Image.open(io.BytesIO(image_bytes))
        image = image.convert('L')
        image = image.resize((image.width * 2, image.height * 2))
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(image, config=custom_config)
        new_leaderboard_data = []
        lines = text.strip().split('\n')

        for line in lines:
            if not line.strip():
                continue

            # Explanation of this pattern -> 
            # ^ - matches start of the line
            # (.*?) - captures the name
            # \s* - optional whitespace
            # (\d{1,3}(?:,\d{3})*) - captures numbers that have commas (score)
            # (\d+ / \d+) - captures tasks done
            # $ - end of the line

            pattern = r'^(.*?)\s*(\d{1,3}(?:,\d{3})*)\s*(\d+/\d+)\s*$'

            match = re.match(pattern, line.strip())
            if match:
                name, score, tasks = match.groups()
                score = int(score.replace(',', ''))
                entry = {
                    "name": name.strip(),
                    "score": score,
                    "tasks": tasks
                }
                new_leaderboard_data.append(entry)

        if new_leaderboard_data:
            final_data = save_data(new_leaderboard_data)
            formatted_data = json.dumps(final_data, indent=2)
            await ctx.send(
                f"Leaderboard data has been processed and saved to data.json!\n```json\n{formatted_data}\n```")
        else:
            await ctx.send("No valid leaderboard entries were found in the image. Please check the debug output above.")

    except Exception as e:
        await ctx.send(f"An error occurred: {str(e)}")
        import traceback
        await ctx.send(f"Debug - Full error:\n```\n{traceback.format_exc()}\n```")

@bot.command(name='getxlsx') # Takes data.josn file and converts it to excel
async def export_to_excel(ctx):
    try:
        if not os.path.exists('data.json'):
            await ctx.send("No data found to export.")
            return

        with open('data.json', 'r', encoding='utf-8') as f:
            data = json.load(f)

        df = pd.DataFrame(data)
        excel_file = 'AM.xlsx'
        df.to_excel(excel_file, index=False)

        await ctx.send("Excel File:", file=discord.File(excel_file))
        os.remove(excel_file)

    except Exception as e:
        await ctx.send(f"An error occurred while exporting to Excel: {str(e)}")

bot.run('YOUR_TOKEN') # Insert your bot's token
