---
name: security-encrypt
description: Encrypt a presentation with password
---

# Encrypt Presentation Example

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// Encrypt with password
presentation.Encrypt("secure_password");

// Or set both open and modify passwords
// presentation.Encrypt("open_password", "modify_password");

presentation.SaveToFile("encrypted.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## Remove Encryption

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("encrypted.pptx", "password");

// Remove password
presentation.RemoveEncryption();

presentation.SaveToFile("unlocked.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```
