using System;
using System.Net.Http;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.InputFiles;
using System.IO;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using TelegramBot;

class Program
{
	private static TelegramBotClient botClient;
	private static readonly HttpClient httpClient = new HttpClient();

	static async Task Main(string[] args)
	{
		var host = Host.CreateDefaultBuilder(args)
			.ConfigureAppConfiguration((context, builder) =>
			{
				// Load appsettings.json
				var currentDirectory = Directory.GetCurrentDirectory();

				// Find appsettings.json starting from the current directory and moving up
				var appSettingsFilePath = FindAppSettingsFile(currentDirectory);

				if (string.IsNullOrEmpty(appSettingsFilePath))
				{
					throw new FileNotFoundException("appsettings.json file was not found.");
				}
								
				builder.AddJsonFile(appSettingsFilePath, optional: false, reloadOnChange: true);
				
			})
			.ConfigureServices((context, services) =>
			{
				// Bind BotSettings to the appsettings.json section

				services.Configure<BotSettings>(context.Configuration.GetSection("BotSettings"));
				// Register the TelegramBot class as a Singleton
				services.AddSingleton<TelegramServiceBot>();
			})
			.Build();

		// Initialize the bot with your token
		var bot = host.Services.GetRequiredService<TelegramServiceBot>();
		await bot.StartBot();
	}

	static string FindAppSettingsFile(string directory)
	{
		// Search for the file in the current directory
		var filePath = Path.Combine(directory, "appsettings.json");

		if (System.IO.File.Exists(filePath))
		{
			return filePath;
		}

		// If not found and there's a parent directory, search in the parent
		var parentDirectory = Directory.GetParent(directory);
		if (parentDirectory != null)
		{
			return FindAppSettingsFile(parentDirectory.FullName);
		}

		// Return null if file is not found
		return null;
	}

}
