import discord
from discord import SyncWebhook


def send_to_discord(webhook_url, output_file):
    """Sending File alerts to discord

    Parameters
    ----------
    webhook_url : string
        _description_
    output_file : string
        _description_
    """
    webhook = SyncWebhook.from_url('https://discordapp.com/api/webhooks/1111659660240486440/oIcLg4rBXdvyxGHAvFvTTg35E-KnPW8C1EXooHIs_y8rjE0IhK061ansb4IRcVzcjCLI')

    with open(file=output_file, mode='rb') as file:
        excel_file = discord.File(file)

    webhook.send('This is an automated report', 
                username='Sales Bot', 
                file=excel_file)