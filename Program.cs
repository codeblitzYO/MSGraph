using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Graph.Auth;

namespace MSGraph
{
	class Program
	{
		static void Main(string[] args)
		{
			_clientApp = PublicClientApplicationBuilder.Create(ClientId)
				.WithAuthority($"{Instance}{Tenant}")
				.WithDefaultRedirectUri()
				.Build();
			TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);

			var program = new Program();
			program.Run();

			while (true) ;
		}

		public async void Run()
		{
			var authProvider = await Login();

			var client = new GraphServiceClient(authProvider);

			//var drives = await client.Drives.Request().GetAsync();
			//foreach (var drive in drives)
			//{
			//	Console.WriteLine(drive.Name);
			//}

			//var mydrive = await client.Me.Drive.Request().GetAsync();

			//var root = await client.Me.Drive.Root.Request().GetAsync();
			//var children = await client.Me.Drive.Root.Children.Request().GetAsync();
			//foreach (var child in children)
			//{
			//	Console.WriteLine(child.Name);
			//}

			var folderItem = new DriveItem
			{
				Name = "uploads",
				Folder = new Folder
				{
				},
				AdditionalData = new Dictionary<string, object>()
				{
					{"@microsoft.graph.conflictBehavior","replace"}
				}
			};
			var folder = await client.Me.Drive.Root.Children.Request().AddAsync(folderItem);
			var filename = System.IO.Path.GetRandomFileName();

			DriveItem uploadedItem = null;
			using (var stream = new System.IO.MemoryStream(new byte[1024 * 10]))
			{
				var uploadSession = await client.Me.Drive.Items[folder.Id].ItemWithPath(filename).CreateUploadSession().Request().PostAsync();

				var maxChunkSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
				var provider = new ChunkedUploadProvider(uploadSession, client, stream, maxChunkSize);

				uploadedItem = await provider.UploadAsync();
			}

			//var me = await client.Me.Request().GetAsync();

			//var url = client.Me.JoinedTeams.RequestUrl;

			//var teams = await client.Me.JoinedTeams.Request().GetAsync();
			//foreach (var team in teams)
			//{
			//	Console.WriteLine(team.Id);
			//}

			//var channels = await client.Teams["9aa92ac0-e14e-4f5c-9748-da145c061844"].Channels.Request().GetAsync();
			//foreach (var channel in channels)
			//{
			//	Console.WriteLine(channel.Id);
			//}

			var fileGuid = Guid.NewGuid();

			var chatMessage = new ChatMessage
			{
				Body = new ItemBody
				{
					Content = "Hello world <attachment id=\"" + fileGuid.ToString() + "\"></attachment>"
				}
			};

			var attachments = new ChatMessageAttachment[]
			{
				new ChatMessageAttachment
				{
					Id = fileGuid.ToString(),
					ContentType = "reference",
					ContentUrl = uploadedItem.WebUrl,
					Name = uploadedItem.Name,
				}
			};
			chatMessage.Attachments = attachments;

			var teamId = "9aa92ac0-e14e-4f5c-9748-da145c061844";
			var channelId = "19:b4803a26abbb442080125c9807084054@thread.tacv2";

			await client.Teams[teamId].Channels[channelId].Messages
				.Request()
				.AddAsync(chatMessage);

			Console.WriteLine(uploadedItem.WebUrl);
		}

		public async Task<IAuthenticationProvider> Login()
		{
			var app = PublicClientApp;
			string[] scopes = new string[] {
				//"Team.ReadBasic.All",
				//"TeamSettings.Read.All",
				//"TeamSettings.ReadWrite.All",
				//"User.Read.All",
				//"User.ReadWrite.All",
				//"Directory.Read.All",
				//"Directory.ReadWrite.All",
				"ChannelMessage.Send",
				//"Group.ReadWrite.All",
				//"Files.Read",
				"Files.ReadWrite",
				//"Files.Read.All", "Files.ReadWrite.All", "Sites.Read.All", "Sites.ReadWrite.All"
			};

			return new InteractiveAuthenticationProvider(app, scopes);

			var accounts = await app.GetAccountsAsync();
			var firstAccount = accounts.FirstOrDefault();

			AuthenticationResult authResult = null;
			try
			{
				authResult = await app.AcquireTokenSilent(scopes, firstAccount)
					.ExecuteAsync();
			}
			catch (MsalUiRequiredException ex)
			{
				// A MsalUiRequiredException happened on AcquireTokenSilent. 
				// This indicates you need to call AcquireTokenInteractive to acquire a token
				System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

				try
				{
					authResult = await app.AcquireTokenInteractive(scopes)
						.WithAccount(accounts.FirstOrDefault())
						//.WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
						.WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
						.ExecuteAsync();
				}
				catch (MsalException msalex)
				{
					Console.WriteLine($"Error Acquiring Token:{System.Environment.NewLine}{msalex}");
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
				return null;
			}

			return null;
		}

		private static string ClientId = "e22f377b-5811-4310-b8fe-39f9d80b759b";

		// Note: Tenant is important for the quickstart. We'd need to check with Andre/Portal if we
		// want to change to the AadAuthorityAudience.
		private static string Tenant = "10ee393f-f114-4bac-8691-8b1a0c214f29";
		private static string Instance = "https://login.microsoftonline.com/";
		private static IPublicClientApplication _clientApp;

		public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
	}
}
