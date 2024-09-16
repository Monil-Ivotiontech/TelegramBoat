using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot.Types.InputFiles;
using Telegram.Bot.Types;
using Telegram.Bot;
using Microsoft.Extensions.Options;
using Telegram.Bot.Polling;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

namespace TelegramBot
{
	public class TelegramServiceBot
	{
		private static BotSettings _botSettings;
		private static TelegramBotClient _botClient;
		private static readonly HttpClient httpClient = new HttpClient();
		public TelegramServiceBot(IOptions<BotSettings> botSettings)
		{
			_botSettings = botSettings.Value;

			// Initialize Telegram Bot Client using the Token from appsettings.json
			_botClient = new TelegramBotClient(_botSettings.Token);
		}
		public async Task StartBot()
		{


			// Set up an event handler to listen for messages
			var cancellationToken = new CancellationTokenSource();

			var receiverOptions = new ReceiverOptions
			{
				AllowedUpdates = Array.Empty<UpdateType>() // receive all update types
			};

			_botClient.StartReceiving(
				updateHandler: HandleUpdateAsync,
				 pollingErrorHandler: HandlePollingErrorAsync,
				receiverOptions: receiverOptions,
				cancellationToken: cancellationToken.Token
			);

			Console.WriteLine("Bot is up and running. Press any key to exit.");
			Console.ReadKey();

			// Stop the bot when the program is exiting
			cancellationToken.Cancel();
		}

		private static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
		{
			var message = update.Message;

			// Check if the message is a command and starts with "/getfile"
			if (message.Text != null)
			{
				if (message.Text.StartsWith("/start"))
				{
					var replyKeyboard = new ReplyKeyboardMarkup(new[]{
					new KeyboardButton[] { "11000", "UPDATE EXCEL" }, // Add more buttons as needed
                });


					// Send a message with the keyboard
					await botClient.SendTextMessageAsync(
						chatId: message.Chat,
						text: "Choose an option:",
						replyMarkup: replyKeyboard
					);
				}
				else
				{
					if (!message.Text.StartsWith("/"))
					{
						var dt = new
						{
							firstName = message.Chat.FirstName ?? "unknown",
							lastName = message.Chat.LastName ?? "",
							fullName = $"{message.Chat.FirstName ?? "unknown"} {message.Chat.LastName ?? ""}",
							userName = message.Chat.LastName ?? "Unknown",
							userId = message.Chat.Id,
						};
						var parameter1 = message.Text;
						var parameter2 = message.Chat.Id;
						// Construct the API URL with the parameters
						var apiUrl = $"{_botSettings.ApiUrl.ToString()}api/file/getfile?parameter1={parameter1}&parameter2={parameter2}";

						// Call your API
						var response = await httpClient.GetAsync(apiUrl);
						if (response.IsSuccessStatusCode)
						{
							var fileBytes = await response.Content.ReadAsByteArrayAsync();

							// Send the file back to the user
							using (var stream = new MemoryStream(fileBytes))
							{
								var fileName = response.Content.Headers.ContentDisposition?.FileName ?? "file";
								await botClient.SendDocumentAsync(message.Chat.Id, new InputOnlineFile(stream, fileName));
							}
						}
						else
						{
							await botClient.SendTextMessageAsync(message.Chat.Id, "Failed to retrieve the file.");
						}
					}
				}
			}

		}

		private static Task HandlePollingErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
		{
			Console.WriteLine($"Polling error: {exception.Message}");
			throw exception;

		}
	}
}
