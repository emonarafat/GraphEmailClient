using System.Net.Http.Headers;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;

/// <summary>
/// Provides methods to interact with Microsoft Graph for email operations.
/// </summary>
namespace GraphEmailClient;

public class EmailClient : IEmailClient
{
    private readonly GraphServiceClient _graphClient; ///< The GraphServiceClient used to make requests to Microsoft Graph.
    private readonly ILogger<EmailClient> _logger; ///< Logger for logging information and errors.

    /// <summary>
    /// Initializes a new instance of the <see cref="EmailClient"/> class.
    /// </summary>
    /// <param name="clientId">The client ID for authentication.</param>
    /// <param name="tenantId">The tenant ID for authentication.</param>
    /// <param name="clientSecret">The client secret for authentication.</param>
    /// <param name="logger">The logger instance for logging.</param>
    /// <exception cref="ArgumentNullException">Thrown when logger is null.</exception>
    public EmailClient(string clientId, string tenantId, string clientSecret, ILogger<EmailClient> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

        var confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantId)
            .WithClientSecret(clientSecret)
            .Build();

        _graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
        {
            var result = await confidentialClientApplication
                .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                .ExecuteAsync();

            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
        }));

        _logger.LogInformation("EmailClient initialized successfully.");
    }

    /// <summary>
    /// Sends an email asynchronously.
    /// </summary>
    /// <param name="subject">The subject of the email.</param>
    /// <param name="body">The body content of the email.</param>
    /// <param name="recipients">A list of recipient email addresses.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="subject"/>, <paramref name="body"/>, or <paramref name="recipients"/> is null or empty.</exception>
    public async Task SendEmailAsync(string subject, string body, List<string> recipients)
    {
        if (string.IsNullOrWhiteSpace(subject) || string.IsNullOrWhiteSpace(body) || recipients == null || !recipients.Any())
        {
            throw new ArgumentNullException("Subject, body, and recipients cannot be null or empty.");
        }

        try
        {
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
                },
                ToRecipients = recipients.Select(email => new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = email
                    }
                }).ToList()
            };

            await _graphClient.Me.SendMail(message, true).Request().PostAsync();
            _logger.LogInformation("Email sent successfully.");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error sending email: {Message}", ex.Message);
            throw; // Rethrow the exception for higher-level handling
        }
    }

    /// <summary>
    /// Reads emails asynchronously.
    /// </summary>
    /// <param name="top">The number of emails to retrieve.</param>
    /// <returns>A task representing the asynchronous operation, with a list of retrieved emails.</returns>
    /// <exception cref="ServiceException">Thrown when there is an error retrieving emails from the service.</exception>
    public async Task<List<Message>> ReadEmailsAsync(int top = 10)
    {
        try
        {
            var messages = await _graphClient.Me.Messages
                .Request()
                .Top(top)
                .GetAsync();

            _logger.LogInformation("Emails retrieved successfully.");
            return messages.CurrentPage.ToList();
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error reading emails: {Message}", ex.Message);
            throw; // Rethrow the exception for higher-level handling
        }
    }

    /// <summary>
    /// Moves an email to a specified folder asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to move.</param>
    /// <param name="destinationFolderId">The ID of the destination folder.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> or <paramref name="destinationFolderId"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error moving the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.MoveEmailAsync("message-id", "folder-id");
    /// </code>
    /// </example>
    public async Task MoveEmailAsync(string messageId, string destinationFolderId)
    {
        try
        {
            await _graphClient.Me.Messages[messageId]
                .Move(destinationFolderId)
                .Request()
                .PostAsync();

            _logger.LogInformation("Email moved successfully.");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error moving email: {Message}",ex.Message);
        }
    }

    /// <summary>
    /// Marks an email as read or unread asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to mark.</param>
    /// <param name="isRead">True to mark as read; false to mark as unread.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error marking the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.MarkEmailAsReadAsync("message-id", true);
    /// </code>
    /// </example>
    public async Task MarkEmailAsReadAsync(string messageId, bool isRead)
    {
        try
        {
            var message = new Message
            {
                IsRead = isRead
            };

            await _graphClient.Me.Messages[messageId]
                .Request()
                .UpdateAsync(message);

            _logger.LogInformation("Email marked as {Status} successfully.", isRead ? "read" : "unread");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error marking email: {Message}",ex.Message);
        }
    }

    /// <summary>
    /// Deletes an email asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to delete.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error deleting the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.DeleteEmailAsync("message-id");
    /// </code>
    /// </example>
    public async Task DeleteEmailAsync(string messageId)
    {
        try
        {
            await _graphClient.Me.Messages[messageId]
                .Request()
                .DeleteAsync();

            _logger.LogInformation("Email deleted successfully.");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error deleting email: {Message}",ex.Message);
        }
    }

    /// <summary>
    /// Lists mail folders asynchronously.
    /// </summary>
    /// <returns>A task representing the asynchronous operation, with a list of mail folders.</returns>
    /// <exception cref="ServiceException">Thrown when there is an error retrieving mail folders.</exception>
    /// <example>
    /// <code>
    /// var folders = await emailClient.ListFoldersAsync();
    /// </code>
    /// </example>
    public async Task<List<MailFolder>> ListFoldersAsync()
    {
        try
        {
            var folders = await _graphClient.Me.MailFolders
                .Request()
                .GetAsync();

            _logger.LogInformation("Mail folders retrieved successfully.");
            return folders.CurrentPage.ToList();
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error listing folders: {Message}", ex.Message);
            return new List<MailFolder>();
        }
    }

    /// <summary>
    /// Creates a new mail folder asynchronously.
    /// </summary>
    /// <param name="folderName">The name of the folder to create.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="folderName"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error creating the folder.</exception>
    /// <example>
    /// <code>
    /// await emailClient.CreateFolderAsync("New Folder");
    /// </code>
    /// </example>
    public async Task CreateFolderAsync(string folderName)
    {
        try
        {
            var mailFolder = new MailFolder
            {
                DisplayName = folderName
            };

            await _graphClient.Me.MailFolders
                .Request()
                .AddAsync(mailFolder);

            _logger.LogInformation("Mail folder created successfully.");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error creating folder: {Message}",ex.Message);
        }
    }

    /// <summary>
    /// Moves an email to the Junk folder asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to move.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error moving the email to the Junk folder.</exception>
    /// <example>
    /// <code>
    /// await emailClient.MoveEmailToJunkAsync("message-id");
    /// </code>
    /// </example>
    public async Task MoveEmailToJunkAsync(string messageId)
    {
        try
        {
            var junkFolder = await _graphClient.Me.MailFolders
                .Request()
                .Filter("displayName eq 'Junk Email'")
                .GetAsync();

            var junkFolderId = junkFolder.CurrentPage.FirstOrDefault()?.Id;

            if (!string.IsNullOrEmpty(junkFolderId))
            {
                await MoveEmailAsync(messageId, junkFolderId);
                _logger.LogInformation("Email moved to junk folder successfully.");
            }
            else
            {
                _logger.LogWarning("Junk folder not found.");
            }
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error moving email to junk: {Message}",ex.Message);
        }
    }

    /// <summary>
    /// Replies to an email asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to reply to.</param>
    /// <param name="replyBody">The body content of the reply.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> or <paramref name="replyBody"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error replying to the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.ReplyToEmailAsync("message-id", "This is a reply.");
    /// </code>
    /// </example>
    public async Task ReplyToEmailAsync(string messageId, string replyBody)
    {
        try
        {
            var reply = new Message
            {
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = replyBody
                }
            };

            await _graphClient.Me.Messages[messageId]
                .CreateReply(reply)
                .Request()
                .PostAsync();

            _logger.LogInformation("Email replied to successfully.");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error replying to email: {Message}",ex.Message);
        }
    }

    /// <summary>
    /// Forwards an email asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to forward.</param>
    /// <param name="forwardBody">The body content of the forwarded email.</param>
    /// <param name="recipients">A list of recipient email addresses for the forwarded email.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/>, <paramref name="forwardBody"/>, or <paramref name="recipients"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error forwarding the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.ForwardEmailAsync("message-id", "Check this out!", new List&lt;string&gt; { "recipient@example.com" });
    /// </code>
    /// </example>
    public async Task ForwardEmailAsync(string messageId, string forwardBody, List<string> recipients)
    {
        try
        {
            var forward = new Message
            {
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = forwardBody
                },
                ToRecipients = recipients.Select(email => new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = email
                    }
                }).ToList()
            };

            await _graphClient.Me.Messages[messageId]
                .Forward(Message: forward)
                .Request()
                .PostAsync();

            _logger.LogInformation("Email forwarded successfully.");
        }
        catch (ServiceException ex)
        {
            _logger.LogError("Error forwarding email: {Message}",ex.Message);
        }
    }

}
