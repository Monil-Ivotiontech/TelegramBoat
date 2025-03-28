
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
        private static CancellationTokenSource cancellationToken;
        public TelegramServiceBot(IOptions<BotSettings> botSettings)
        {
            _botSettings = botSettings.Value;

            // Initialize Telegram Bot Client using the Token from appsettings.json
            _botClient = new TelegramBotClient(_botSettings.Token);
            cancellationToken = new CancellationTokenSource();
        }
        public async Task StartBot()
        {
            // Set up an event handler to listen for messages

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
            if (message != null && message.Text != null)
            {
                if (message.Text.StartsWith("/start"))
                {
                    var replyKeyboard = new ReplyKeyboardMarkup(new[]{
                    new KeyboardButton[] { "UPDATE ALL EXCEL", "COM POS EXCEL", "M2M EXCEL" }, // Add more buttons as needed
                });


                    // Send a message with the keyboard
                    await botClient.SendTextMessageAsync(
                        chatId: message.Chat,
                        text: "Wellcome to Pdf bot, Now you can start"
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
                        var parameter2 = message.Chat.Id.ToString();
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
                    else {
                        await botClient.SendTextMessageAsync(
                           chatId: message.Chat,
                           text: "Message is not in a correct format!"
                       );
                    }
                }
            }

        }

        private static async Task HandlePollingErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            Console.WriteLine($"Polling error: {exception.Message}");

            // Log the error and restart the bot
            await Task.Delay(5000); // Wait for 5 seconds before restarting
            Console.WriteLine("Restarting bot...");

            await RestartBot();

        }


        private static async Task RestartBot()
        {
            try
            {
                // Cancel existing token and create a new one
                cancellationToken.Cancel();
                cancellationToken = new CancellationTokenSource();

                var receiverOptions = new ReceiverOptions
                {
                    AllowedUpdates = Array.Empty<UpdateType>() // Receive all update types
                };

                _botClient.StartReceiving(
                    updateHandler: HandleUpdateAsync,
                    pollingErrorHandler: HandlePollingErrorAsync,
                    receiverOptions: receiverOptions,
                    cancellationToken: cancellationToken.Token
                );

                Console.WriteLine("Bot restarted successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while restarting bot: {ex.Message}");
                await Task.Delay(5000);
                await RestartBot(); // Retry restarting after 5 seconds
            }
        }

    }
}
