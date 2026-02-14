# Bulk Mail Sender Application

A powerful desktop application for sending bulk emails using Microsoft Outlook with MongoDB database integration.

## Features

### 1. **Drafts Management**
- Create, read, update, and delete email drafts
- Integration with Microsoft Outlook drafts
- Tab-based view for Application and Outlook drafts
- Open drafts directly in Outlook
- Sync drafts from Outlook to application database

### 2. **Send Mail**
- Select recipients by industry
- Import recipients from text files
- Batch email sending with configurable batch size
- BCC field support (To field kept empty)
- Configurable time delay between batches
- Mark emails as sent/unsent
- Real-time progress tracking
- Unique whiteline text in each batch for tracking
- Preview drafts before sending

### 3. **Recipient Management**
- Industry-wise recipient organization
- Single email can belong to multiple industries
- Import recipients from text files
- Automatic duplicate handling (merges industries)
- Search and filter capabilities
- Mark recipients as sent/unsent

### 4. **Inbox & Outbox**
- View Outlook inbox messages
- View sent items from Outlook
- Open emails directly in Outlook

### 5. **Email Accounts**
- Manage multiple email accounts
- Set default account for sending
- Account information storage

### 6. **Settings**
- Industry management (create, edit, delete)
- Configure default batch size
- Configure default delay between batches
- Application-wide settings

## Technical Stack

- **.NET 8.0** - Windows Forms Application
- **MongoDB** - Database for storing recipients, drafts, industries, and settings
- **Microsoft Office Interop (Outlook)** - Email sending and Outlook integration
- **Modern UI Design** - Clean white theme with responsive layout

## Prerequisites

1. .NET 8.0 SDK
2. MongoDB Server (running on localhost:27017 or configured connection string)
3. Microsoft Outlook installed and configured

## Installation

1. Clone or download the project
2. Ensure MongoDB is running
3. Open the solution in Visual Studio 2022 or later
4. Restore NuGet packages
5. Build and run the application

## Database Configuration

Default MongoDB connection:
- **Connection String**: `mongodb://localhost:27017`
- **Database Name**: `BulkMailSender`

Collections:
- `recipients` - Email recipients with industry associations
- `drafts` - Email templates
- `industries` - Industry categories
- `emailAccounts` - Email account information
- `settings` - Application settings

## Usage

### First Time Setup

1. **Add Industries**: Go to Settings tab and create industries (e.g., Technology, Healthcare, Finance)
2. **Import Recipients**: Use the Recipient List tab to import emails from a text file
3. **Create Drafts**: Create email templates in the Drafts tab
4. **Configure Settings**: Set default batch size and delay in Settings

### Sending Bulk Emails

1. Navigate to **Send Mail** tab
2. Select an industry to filter recipients
3. Select a draft template
4. Configure batch size (number of emails in BCC per batch)
5. Set delay between batches (in seconds)
6. Check recipients you want to send to
7. Click **Start Sending**

### Importing Recipients

Text file format (one email per line):
```
john@example.com
jane@company.com
contact@business.org
```

The application will:
- Remove duplicates
- Validate email format
- Associate with selected industries
- Merge industries if email already exists

## Features in Detail

### Batch Processing
- Emails are sent in batches using BCC field
- To field is kept empty for privacy
- Each batch includes a unique invisible whiteline text for tracking
- Configurable delay between batches to avoid spam filters

### Industry Management
- Recipients can belong to multiple industries
- Filter and send emails by industry
- Easy industry assignment during import

### Outlook Integration
- Seamless integration with Microsoft Outlook
- Use Outlook's sending infrastructure
- Access Outlook drafts, inbox, and sent items
- Open emails directly in Outlook

## UI Design

- **Left Sidebar**: Navigation menu with 7 main sections
- **Right Content Area**: Dynamic content based on selected menu
- **Responsive Design**: Scrollable content areas
- **Modern Theme**: Clean white design with blue accents
- **Premium Look**: Professional admin panel interface

## Security Notes

- Email credentials are managed through Outlook
- No password storage in the application
- MongoDB connection should be secured in production
- Consider using MongoDB authentication

## Troubleshooting

**Outlook not connecting:**
- Ensure Outlook is installed and configured
- Run the application with administrator privileges if needed

**MongoDB connection failed:**
- Verify MongoDB is running
- Check connection string in MongoDbService.cs
- Ensure MongoDB port (27017) is not blocked

**Emails not sending:**
- Check Outlook is properly configured
- Verify email account is set up in Outlook
- Check internet connection

## Future Enhancements

- Email scheduling
- Email templates with variables
- Detailed sending reports
- Email open tracking
- Attachment support
- SMTP direct sending option
- Multi-language support

## License

This is a proprietary application. All rights reserved.

## Support

For issues or questions, please contact the development team.
