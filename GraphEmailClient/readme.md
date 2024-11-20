# GraphEmailClient

GraphEmailClient is a .NET library that provides methods to interact with Microsoft Graph for email operations. This library allows you to send, read, move, and manage emails seamlessly using the Microsoft Graph API.

## Features

- Send emails asynchronously
- Read emails asynchronously
- Move emails to specified folders
- Mark emails as read or unread
- Delete emails
- List and create mail folders
- Move emails to the Junk folder
- Reply to and forward emails

## Installation

You can install the GraphEmailClient NuGet package using one of the following methods:

### .NET CLI

Run the following command in your terminal:

```
dotnet add package GraphEmailClient
```

### NuGet Package Manager Console

If you prefer using the NuGet Package Manager Console, run the following command:

```
Install-Package GraphEmailClient
```

## Usage

Here’s a quick example of how to use the GraphEmailClient:

```csharp
using Microsoft.Extensions.Logging;
using GraphEmailClient;

var loggerFactory = LoggerFactory.Create(builder => {
    builder.AddConsole();
});
ILogger<EmailClient> logger = loggerFactory.CreateLogger<EmailClient>();

var emailClient = new EmailClient("clientId", "tenantId", "clientSecret", logger);

// Send an email
await emailClient.SendEmailAsync("Hello", "This is a test email.", new List<string> { "example@example.com" });

// Read emails
var emails = await emailClient.ReadEmailsAsync(5);

// Move an email
await emailClient.MoveEmailAsync("message-id", "folder-id");

// Mark an email as read
await emailClient.MarkEmailAsReadAsync("message-id", true);

// Delete an email
await emailClient.DeleteEmailAsync("message-id");

// List mail folders
var folders = await emailClient.ListFoldersAsync();

// Create a new folder
await emailClient.CreateFolderAsync("New Folder");

// Move an email to Junk
await emailClient.MoveEmailToJunkAsync("message-id");

// Reply to an email
await emailClient.ReplyToEmailAsync("message-id", "This is a reply.");

// Forward an email
await emailClient.ForwardEmailAsync("message-id", "Check this out!", new List<string> { "recipient@example.com" });
```

## Method Descriptions

### SendEmailAsync

- **Description**: Sends an email asynchronously.
- **Parameters**:
  - `subject`: The subject of the email.
  - `body`: The body content of the email.
  - `recipients`: A list of recipient email addresses.
- **Exceptions**: Throws `ArgumentNullException` if any parameter is null or empty.

### ReadEmailsAsync

- **Description**: Reads emails asynchronously.
- **Parameters**:
  - `top`: The number of emails to retrieve (default is 10).
- **Returns**: A list of retrieved emails.
- **Exceptions**: Throws `ServiceException` if there is an error retrieving emails.

### MoveEmailAsync

- **Description**: Moves an email to a specified folder asynchronously.
- **Parameters**:
  - `messageId`: The ID of the email to move.
  - `destinationFolderId`: The ID of the destination folder.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException` if parameters are invalid.

### MarkEmailAsReadAsync

- **Description**: Marks an email as read or unread asynchronously.
- **Parameters**:
  - `messageId`: The ID of the email to mark.
  - `isRead`: True to mark as read; false to mark as unread.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException`.

### DeleteEmailAsync

- **Description**: Deletes an email asynchronously.
- **Parameters**:
  - `messageId`: The ID of the email to delete.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException`.

### ListFoldersAsync

- **Description**: Lists mail folders asynchronously.
- **Returns**: A list of mail folders.
- **Exceptions**: Throws `ServiceException`.

### CreateFolderAsync

- **Description**: Creates a new mail folder asynchronously.
- **Parameters**:
  - `folderName`: The name of the folder to create.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException`.

### MoveEmailToJunkAsync

- **Description**: Moves an email to the Junk folder asynchronously.
- **Parameters**:
  - `messageId`: The ID of the email to move.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException`.

### ReplyToEmailAsync

- **Description**: Replies to an email asynchronously.
- **Parameters**:
  - `messageId`: The ID of the email to reply to.
  - `replyBody`: The body content of the reply.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException`.

### ForwardEmailAsync

- **Description**: Forwards an email asynchronously.
- **Parameters**:
  - `messageId`: The ID of the email to forward.
  - `forwardBody`: The body content of the forwarded email.
  - `recipients`: A list of recipient email addresses for the forwarded email.
- **Exceptions**: Throws `ArgumentNullException` or `ServiceException`.

## Documentation

For detailed documentation on the available methods and their usage, please refer to the XML comments in the code. Each method is documented with its parameters, return types, and exceptions. You can also find usage examples in the **Usage** section above.

### XML Documentation

The library includes XML documentation comments that provide additional context and examples for each method. You can view this documentation directly in your IDE or by generating documentation files.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.

## Author

Created by [Yaseer Arafat](https://github.com/emoarafat).
