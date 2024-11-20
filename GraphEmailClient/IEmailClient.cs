using Microsoft.Graph;

/// <summary>
/// Defines the contract for email operations using Microsoft Graph.
/// </summary>
namespace GraphEmailClient;

public interface IEmailClient
{
    /// <summary>
    /// Sends an email asynchronously.
    /// </summary>
    /// <param name="subject">The subject of the email.</param>
    /// <param name="body">The body content of the email.</param>
    /// <param name="recipients">A list of recipient email addresses.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="subject"/>, <paramref name="body"/>, or <paramref name="recipients"/> is null or empty.</exception>
    /// <example>
    /// <code>
    /// await emailClient.SendEmailAsync("Hello", "This is a test email.", new List&lt;string&gt; { "example@example.com" });
    /// </code>
    /// </example>
    Task SendEmailAsync(string subject, string body, List<string> recipients);

    /// <summary>
    /// Reads emails asynchronously.
    /// </summary>
    /// <param name="top">The number of emails to retrieve.</param>
    /// <returns>A task representing the asynchronous operation, with a list of retrieved emails.</returns>
    /// <exception cref="ServiceException">Thrown when there is an error retrieving emails from the service.</exception>
    /// <example>
    /// <code>
    /// var emails = await emailClient.ReadEmailsAsync(5);
    /// </code>
    /// </example>
    Task<List<Message>> ReadEmailsAsync(int top = 10);

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
    Task MoveEmailAsync(string messageId, string destinationFolderId);

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
    Task MarkEmailAsReadAsync(string messageId, bool isRead);

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
    Task DeleteEmailAsync(string messageId);

    /// <summary>
    /// Lists mail folders asynchronously.
    /// </summary>
    /// <returns>A task representing the asynchronous operation, with a list of mail folders.</returns>
    /// <exception cref="ServiceException">Thrown when there is an error listing mail folders.</exception>
    /// <example>
    /// <code>
    /// var folders = await emailClient.ListFoldersAsync();
    /// </code>
    /// </example>
    Task<List<MailFolder>> ListFoldersAsync();

    /// <summary>
    /// Creates a new mail folder asynchronously.
    /// </summary>
    /// <param name="folderName">The name of the folder to create.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="folderName"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error creating the folder.</exception>
    /// <example>
    /// <code>
    /// await emailClient.CreateFolderAsync("Inbox");
    /// </code>
    /// </example>
    Task CreateFolderAsync(string folderName);

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
    Task MoveEmailToJunkAsync(string messageId);

    /// <summary>
    /// Replies to an email asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to reply to.</param>
    /// <param name="replyBody">The body content of the reply.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error replying to the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.ReplyToEmailAsync("message-id", "This is a reply.");
    /// </code>
    /// </example>
    Task ReplyToEmailAsync(string messageId, string replyBody);

    /// <summary>
    /// Forwards an email asynchronously.
    /// </summary>
    /// <param name="messageId">The ID of the email to forward.</param>
    /// <param name="forwardBody">The body content of the forwarded email.</param>
    /// <param name="recipients">A list of recipient email addresses for the forwarded email.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="messageId"/> is null or empty.</exception>
    /// <exception cref="ServiceException">Thrown when there is an error forwarding the email.</exception>
    /// <example>
    /// <code>
    /// await emailClient.ForwardEmailAsync("message-id", "This is a forwarded email.", new List&lt;string&gt; { "forwarded@example.com" });
    /// </code>
    /// </example>
    Task ForwardEmailAsync(string messageId, string forwardBody, List<string> recipients);
}