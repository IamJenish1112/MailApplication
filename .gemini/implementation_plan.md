# HTML Email Template Designer - Architecture & Implementation Plan

## 📋 Overview

Replace the legacy `WebBrowser` control (IE-based) with **WebView2** (Chromium-based) and embed **TinyMCE 6** as the professional rich text HTML editor. This eliminates all broken styling issues caused by the outdated IE rendering engine.

---

## 🏗️ Architecture

```
┌──────────────────────────────────────────────────────────┐
│                    MainForm (WinForms)                    │
│  ┌─────────┐  ┌──────────────────────────────────────┐   │
│  │ Sidebar  │  │         Content Panel                │   │
│  │          │  │  ┌──────────────────────────────┐    │   │
│  │ Drafts   │  │  │   DraftEditorForm            │    │   │
│  │ Send     │  │  │  ┌────────────────────────┐  │    │   │
│  │ Inbox    │  │  │  │  WebView2 Control       │  │    │   │
│  │ Outbox   │  │  │  │  ┌──────────────────┐  │  │    │   │
│  │ ...      │  │  │  │  │  TinyMCE Editor   │  │  │    │   │
│  │          │  │  │  │  │  (HTML/JS)         │  │  │    │   │
│  │          │  │  │  │  └──────────────────┘  │  │    │   │
│  │          │  │  │  └────────────────────────┘  │    │   │
│  │          │  │  └──────────────────────────────┘    │   │
│  └─────────┘  └──────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────┘
```

### Service Layer

```
Services/
├── MongoDbService.cs          (existing - add EmailTemplate collection)
├── EditorBridgeService.cs     (NEW - JS ↔ C# communication)
├── DraftDesignerService.cs    (NEW - template management logic)
├── TemplateRepository.cs      (NEW - MongoDB CRUD for templates)
├── ImageStorageService.cs     (NEW - image handling: base64/disk/GridFS)
└── HtmlSanitizer.cs           (NEW - clean & inline CSS for email)
```

### Models

```
Models/
├── Draft.cs                   (existing - add TemplateCategoryId)
├── EmailTemplate.cs           (NEW - reusable template model)
└── TemplateImage.cs           (NEW - image metadata model)
```

---

## 🔧 Technology Choices

| Component | Choice | Reason |
|-----------|--------|--------|
| Browser Control | **WebView2** | Chromium-based, modern CSS/JS support |
| Rich Text Editor | **TinyMCE 6 (Self-hosted)** | Best email-compatible HTML output, free community version, excellent table/image support |
| Image Storage | **Base64 inline** (primary) + **disk path** (optional) | Email compatibility; base64 works in all email clients |
| HTML Cleanup | **PreMailer.Net** NuGet | Converts CSS to inline styles for email clients |

### Why TinyMCE over alternatives?
- **Quill**: Outputs `<div>` based HTML, poor table support, not email-optimized
- **CKEditor**: Requires license for commercial use
- **TinyMCE**: Free community tier, outputs clean `<table>`-based HTML, has email-specific plugins

---

## 📦 NuGet Packages Required

```xml
<PackageReference Include="Microsoft.Web.WebView2" Version="1.0.2903.40" />
<PackageReference Include="PreMailer.Net" Version="2.5.0" />
```

---

## 📁 File Structure (New/Modified Files)

```
MailApplication/
├── Models/
│   ├── Draft.cs                    (MODIFIED)
│   ├── EmailTemplate.cs           (NEW)
│   └── TemplateImage.cs           (NEW)
├── Services/
│   ├── MongoDbService.cs          (MODIFIED)
│   ├── EditorBridgeService.cs     (NEW)
│   ├── DraftDesignerService.cs    (NEW)
│   ├── TemplateRepository.cs      (NEW)
│   ├── ImageStorageService.cs     (NEW)
│   └── HtmlSanitizerService.cs    (NEW)
├── Forms/
│   ├── DraftEditorForm.cs         (REWRITTEN)
│   ├── DraftPreviewForm.cs        (MODIFIED)
│   └── HtmlSourceDialog.cs        (NEW)
├── wwwroot/
│   └── editor/
│       └── editor.html            (NEW - TinyMCE host page)
└── MailApplication.csproj         (MODIFIED)
```

---

## 📐 MongoDB Schema

### EmailTemplate Collection
```json
{
  "_id": ObjectId,
  "name": "Welcome Email",
  "category": "onboarding",
  "subject": "Welcome {{Name}}!",
  "htmlBody": "<html>...</html>",
  "rawEditorHtml": "<html>...</html>",
  "placeholders": ["{{Name}}", "{{Company}}", "{{OrderId}}"],
  "thumbnailBase64": "data:image/png;base64,...",
  "isActive": true,
  "createdAt": ISODate,
  "updatedAt": ISODate
}
```

### TemplateImage Collection
```json
{
  "_id": ObjectId,
  "templateId": "ref_to_template",
  "fileName": "logo.png",
  "mimeType": "image/png",
  "base64Data": "iVBORw0KGgo...",
  "filePath": "C:\\Images\\logo.png",
  "storageType": "base64",
  "createdAt": ISODate
}
```

---

## 🔌 JS ↔ C# Communication (EditorBridgeService)

### C# → JavaScript (Setting content)
```csharp
await webView.CoreWebView2.ExecuteScriptAsync(
    $"tinymce.activeEditor.setContent(`{escapedHtml}`)");
```

### JavaScript → C# (Getting content)
```csharp
string html = await webView.CoreWebView2.ExecuteScriptAsync(
    "tinymce.activeEditor.getContent()");
```

### C# Host Object (for callbacks from JS)
```csharp
[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]
public class EditorBridgeService
{
    public event Action<string>? OnContentChanged;
    public event Action<string>? OnImageInsertRequested;
    
    public void NotifyContentChanged(string html) => OnContentChanged?.Invoke(html);
    public void RequestImageInsert(string context) => OnImageInsertRequested?.Invoke(context);
}
```

---

## 📝 Implementation Steps

### Step 1: Install NuGet Packages
```bash
dotnet add package Microsoft.Web.WebView2
dotnet add package PreMailer.Net
```

### Step 2: Download TinyMCE
Download TinyMCE 6 Community from https://www.tiny.cloud/get-tiny/self-hosted/
Extract to `wwwroot/editor/tinymce/`

### Step 3: Create editor.html
The embedded HTML page that hosts TinyMCE with all toolbar options.

### Step 4: Create new service files
- EditorBridgeService.cs
- DraftDesignerService.cs  
- TemplateRepository.cs
- ImageStorageService.cs
- HtmlSanitizerService.cs

### Step 5: Rewrite DraftEditorForm
Replace WebBrowser with WebView2, wire up EditorBridgeService.

### Step 6: Update models and MongoDbService

### Step 7: Add template preview with WebView2

---

## ⚠️ Error Handling Strategy

| Scenario | Handling |
|----------|----------|
| Editor not loaded | Show loading spinner, retry 3 times, then show fallback textarea |
| WebView2 crash | Catch `CoreWebView2ProcessFailed`, restart WebView2, restore content from auto-save |
| Invalid HTML | Run through HtmlSanitizerService before save, show warning if malformed |
| Image too large | Warn if > 500KB, compress or reject > 2MB |
| MongoDB connection failure | Queue saves locally, retry on reconnect |

---

## 🎨 Dynamic Placeholders

Support these placeholders in templates:
- `{{Name}}` - Recipient name
- `{{Company}}` - Company name
- `{{OrderId}}` - Order identifier
- `{{Email}}` - Recipient email
- `{{Date}}` - Current date
- `{{UnsubscribeLink}}` - Unsubscribe URL

Placeholders are highlighted in the editor with a custom TinyMCE plugin using styled `<span>` tags.

---

## 🖼️ Image Handling Strategy

1. **Base64 Embedding** (Default for email): Convert image to base64, embed directly in `<img src="data:...">`
2. **Disk Storage**: Save to `AppData/MailApplication/Images/`, reference by path (for large images)
3. **MongoDB GridFS**: For template images that need to persist across machines

The `ImageStorageService` handles all three strategies with a configurable default.
