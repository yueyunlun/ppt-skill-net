---
title: 安全性
category: spire-presentation
description: 使用 Spire.Presentation 进行文档加密、解密、数字签名等安全操作
---

# 安全性

## 概述

Spire.Presentation 提供了完整的安全功能来保护演示文稿，包括：
- 文档加密和解密
- 设置打开和修改密码
- 添加数字签名
- 移除数字签名
- 设置文档为只读

## 文档加密

### 示例 1: 加密文档

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 设置打开密码（只读密码）
presentation.Encrypt("open_password");

// 或同时设置打开密码和修改密码
presentation.Encrypt("open_password", "modify_password");

presentation.SaveToFile("Encrypted.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 2: 检查文档是否受密码保护

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 尝试不使用密码加载文件
try
{
    presentation.LoadFromFile("document.pptx");
    Console.WriteLine("文档未受密码保护");
}
catch (Exception)
{
    Console.WriteLine("文档受密码保护");
}
finally
{
    presentation.Dispose();
}
```

### 示例 3: 打开加密文档

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 使用密码加载加密文档
presentation.LoadFromFile("encrypted.pptx", "open_password");

Console.WriteLine("文档成功打开");

presentation.SaveToFile("unlocked.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 4: 移除密码保护

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 加载加密文档
presentation.LoadFromFile("encrypted.pptx", "password");

// 移除密码
presentation.RemoveEncryption();

// 保存无密码版本
presentation.SaveToFile("unlocked.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 5: 修改密码

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();

// 加载使用旧密码的文档
presentation.LoadFromFile("protected.pptx", "old_password");

// 设置新密码
presentation.Encrypt("new_password");

presentation.SaveToFile("reprotected.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 数字签名

### 示例 6: 添加数字签名

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 添加数字签名
presentation.AddDigitalSignature(
    "certificate.pfx",      // 证书文件路径
    "certificate_password", // 证书密码
    "签名者姓名",           // 签名者名称
    "这是文档签名"           // 签名描述
);

presentation.SaveToFile("Signed.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 7: 添加带日期的数字签名

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;
using System;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 添加数字签名（自动包含日期）
presentation.AddDigitalSignature(
    "certificate.pfx",
    "certificate_password",
    "签名者",
    $"文档签名 - {DateTime.Now:yyyy-MM-dd}"
);

presentation.SaveToFile("SignedWithDate.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 8: 检查文档数字签名

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;

Presentation presentation = new Presentation();
presentation.LoadFromFile("document.pptx");

// 检查是否有数字签名
if (presentation.IsDigitallySigned)
{
    Console.WriteLine("文档包含数字签名");

    // 获取数字签名信息
    DigitalSignatureCollection signatures = presentation.GetDigitalSignatures();

    foreach (DigitalSignature signature in signatures)
    {
        Console.WriteLine($"签名者: {signature.SignerName}");
        Console.WriteLine($"时间: {signature.SignTime}");
        Console.WriteLine($"描述: {signature.Comments}");
        Console.WriteLine($"证书主题: {signature.Certificate.Subject}");
        Console.WriteLine($"是否有效: {signature.IsValid}");
    }
}
else
{
    Console.WriteLine("文档未签名");
}

presentation.Dispose();
```

### 示例 9: 移除所有数字签名

```csharp
using Spire.Presentation;

Presentation presentation = new Presentation();
presentation.LoadFromFile("signed.pptx");

// 移除所有数字签名
presentation.RemoveAllDigitalSignatures();

presentation.SaveToFile("Unsigned.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 10: 检查数字签名有效性

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;

Presentation presentation = new Presentation();
presentation.LoadFromFile("document.pptx");

if (presentation.IsDigitallySigned)
{
    DigitalSignatureCollection signatures = presentation.GetDigitalSignatures();

    bool allValid = true;
    foreach (DigitalSignature signature in signatures)
    {
        if (!signature.IsValid)
        {
            Console.WriteLine($"签名 {signature.SignerName} 无效！");
            allValid = false;
        }
        else
        {
            Console.WriteLine($"签名 {signature.SignerName} 有效");
        }
    }

    if (allValid)
    {
        Console.WriteLine("所有签名均有效");
    }
}

presentation.Dispose();
```

## 只读保护

### 示例 11: 设置文档为只读

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;

Presentation presentation = new Presentation();
presentation.LoadFromFile("input.pptx");

// 将文档标记为最终状态（只读）
presentation.MarkDocumentAsFinal();

presentation.SaveToFile("ReadOnly.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

### 示例 12: 检查文档是否为只读

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;

Presentation presentation = new Presentation();
presentation.LoadFromFile("document.pptx");

// 检查是否为只读
if (presentation.IsFinalized)
{
    Console.WriteLine("文档标记为最终状态（只读）");
}
else
{
    Console.WriteLine("文档可以编辑");
}

presentation.Dispose();
```

### 示例 13: 移除只读标记

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;

Presentation presentation = new Presentation();
presentation.LoadFromFile("readonly.pptx");

// 移除最终状态标记
presentation.RemoveFinalState();

presentation.SaveToFile("Editable.pptx", FileFormat.Pptx2010);
presentation.Dispose();
```

## 综合安全操作

### 示例 14: 完整的安全保护流程

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;
using System;

Presentation presentation = new Presentation();
presentation.LoadFromFile("confidential.pptx");

// 1. 添加水印保护
foreach (ISlide slide in presentation.Slides)
{
    // 添加水印代码...
}

// 2. 设置密码保护
presentation.Encrypt("confidential_password");

// 3. 添加数字签名
presentation.AddDigitalSignature(
    "certificate.pfx",
    "cert_password",
    "安全管理员",
    "机密文档 - 请勿外传"
);

// 4. 标记为只读
presentation.MarkDocumentAsFinal();

// 5. 保存受保护的文档
presentation.SaveToFile("FullyProtected.pptx", FileFormat.Pptx2010);
presentation.Dispose();

Console.WriteLine("文档已完全保护：");
Console.WriteLine("- 密码保护已启用");
Console.WriteLine("- 数字签名已添加");
Console.WriteLine("- 标记为只读");
```

### 示例 15: 批量处理文档安全

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;
using System.IO;

// 获取目录下所有 PPT 文件
string[] pptFiles = Directory.GetFiles("input_folder", "*.pptx");

foreach (string file in pptFiles)
{
    using (Presentation presentation = new Presentation())
    {
        presentation.LoadFromFile(file);

        // 添加统一密码
        presentation.Encrypt("company_password");

        // 保存到输出目录
        string outputFile = Path.Combine("output_folder", Path.GetFileName(file));
        presentation.SaveToFile(outputFile, FileFormat.Pptx2010);
    }

    Console.WriteLine($"已处理: {file}");
}
```

### 示例 16: 安全检查报告

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;
using System;

Presentation presentation = new Presentation();

try
{
    // 尝试不使用密码加载
    presentation.LoadFromFile("document.pptx");

    Console.WriteLine("=== 文档安全报告 ===");
    Console.WriteLine($"文件名: document.pptx");
    Console.WriteLine($"密码保护: 否");
    Console.WriteLine($"数字签名: {presentation.IsDigitallySigned}");
    Console.WriteLine($"只读状态: {presentation.IsFinalized}");
}
catch
{
    // 文件可能受密码保护
    Console.WriteLine("=== 文档安全报告 ===");
    Console.WriteLine($"文件名: document.pptx");
    Console.WriteLine($"密码保护: 是（需要密码才能访问）");
}
finally
{
    presentation.Dispose();
}
```

### 示例 17: 密码强度验证

```csharp
using System;
using System.Text.RegularExpressions;

// 验证密码强度
PasswordStrength ValidatePassword(string password)
{
    if (string.IsNullOrEmpty(password))
        return PasswordStrength.Weak;

    int score = 0;

    // 长度检查
    if (password.Length >= 8) score++;
    if (password.Length >= 12) score++;

    // 复杂性检查
    if (Regex.IsMatch(password, @"[a-z]")) score++;       // 小写字母
    if (Regex.IsMatch(password, @"[A-Z]")) score++;       // 大写字母
    if (Regex.IsMatch(password, @"[0-9]")) score++;       // 数字
    if (Regex.IsMatch(password, @"[!@#$%^&*(),.?""{}|<>]")) score++; // 特殊字符

    if (score < 4) return PasswordStrength.Weak;
    if (score < 6) return PasswordStrength.Medium;
    return PasswordStrength.Strong;
}

enum PasswordStrength
{
    Weak,
    Medium,
    Strong
}

// 使用示例
string password = "MySecure123!";
PasswordStrength strength = ValidatePassword(password);
Console.WriteLine($"密码强度: {strength}");
```

### 示例 18: 自动化文档保护流程

```csharp
using Spire.Presentation;
using Spire.Presentation.Security;
using System;
using System.IO;

// 保护文档类
class DocumentProtector
{
    private string certificatePath;
    private string certificatePassword;

    public DocumentProtector(string certPath, string certPassword)
    {
        certificatePath = certPath;
        certificatePassword = certPassword;
    }

    public void ProtectDocument(string inputFile, string outputFile,
        string openPassword, string signerName, string comment)
    {
        using (Presentation presentation = new Presentation())
        {
            Console.WriteLine($"正在处理: {inputFile}");

            // 加载文档
            presentation.LoadFromFile(inputFile);

            // 添加密码保护
            if (!string.IsNullOrEmpty(openPassword))
            {
                presentation.Encrypt(openPassword);
                Console.WriteLine("  ✓ 密码保护已添加");
            }

            // 添加数字签名
            if (!string.IsNullOrEmpty(certificatePath))
            {
                presentation.AddDigitalSignature(
                    certificatePath,
                    certificatePassword,
                    signerName,
                    comment
                );
                Console.WriteLine("  ✓ 数字签名已添加");
            }

            // 保存
            presentation.SaveToFile(outputFile, FileFormat.Pptx2010);
            Console.WriteLine($"  ✓ 已保存到: {outputFile}");
        }
    }
}

// 使用示例
var protector = new DocumentProtector("cert.pfx", "cert_pass");
protector.ProtectDocument(
    "input.pptx",
    "protected.pptx",
    "secure123!",
    "安全管理员",
    "自动保护文档"
);
```

## 安全最佳实践

### 示例 19: 安全配置管理

```csharp
using System;

// 安全配置类
class SecurityConfig
{
    public string CertificatePath { get; set; }
    public string CertificatePassword { get; set; }
    public string DefaultPassword { get; set; }
    public bool RequireDigitalSignature { get; set; }
    public bool MarkAsFinal { get; set; }

    public static SecurityConfig LoadDefault()
    {
        return new SecurityConfig
        {
            CertificatePath = "cert.pfx",
            CertificatePassword = "password",
            DefaultPassword = "default_secure_pass",
            RequireDigitalSignature = true,
            MarkAsFinal = true
        };
    }
}
```

### 示例 20: 安全审计日志

```csharp
using System;
using System.IO;

// 记录安全操作
class SecurityLogger
{
    private string logFile;

    public SecurityLogger(string logFilePath)
    {
        logFile = logFilePath;
    }

    public void LogOperation(string operation, string document, string user, string details = "")
    {
        string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
                         $"Operation: {operation}, " +
                         $"Document: {document}, " +
                         $"User: {user}, " +
                         $"Details: {details}";

        File.AppendAllText(logFile, logEntry + Environment.NewLine);
    }

    public void LogPasswordSet(string document, string user, PasswordStrength strength)
    {
        LogOperation("Password Set", document, user, $"Strength: {strength}");
    }

    public void LogSignatureAdded(string document, string user, string signer)
    {
        LogOperation("Digital Signature Added", document, user, $"Signer: {signer}");
    }
}

// 使用示例
var logger = new SecurityLogger("security.log");
logger.LogPasswordSet("document.pptx", "admin", PasswordStrength.Strong);
logger.LogSignatureAdded("document.pptx", "admin", "证书持有人");
```

## 安全相关类和属性

### Presentation 安全方法

| 方法 | 描述 |
|------|------|
| `Encrypt(password)` | 设置密码保护 |
| `Encrypt(openPwd, modifyPwd)` | 设置打开和修改密码 |
| `RemoveEncryption()` | 移除密码保护 |
| `AddDigitalSignature(certPath, certPwd, name, comment)` | 添加数字签名 |
| `RemoveAllDigitalSignatures()` | 移除所有数字签名 |
| `MarkDocumentAsFinal()` | 标记为只读 |
| `RemoveFinalState()` | 移除只读标记 |

### Presentation 安全属性

| 属性 | 描述 |
|------|------|
| `IsDigitallySigned` | 是否包含数字签名 |
| `IsFinalized` | 是否标记为最终状态 |
| `GetDigitalSignatures()` | 获取数字签名集合 |

### DigitalSignature 属性

| 属性 | 描述 |
|------|------|
| `SignerName` | 签名者姓名 |
| `SignTime` | 签名时间 |
| `Comments` | 签名描述 |
| `Certificate` | 数字证书 |
| `IsValid` | 签名是否有效 |

## 注意事项

1. **密码复杂性**: 使用强密码（包含大小写字母、数字和特殊字符）
2. **证书管理**: 妥善保管数字证书和密码
3. **密码恢复**: Spire.Presentation 不支持密码恢复，忘记密码可能导致无法打开文档
4. **版本兼容性**: 某些安全功能在不同版本中可能表现不同

## 最佳实践

1. **分层保护**: 结合使用密码和数字签名
2. **定期更新**: 定期更改密码和更新证书
3. **备份策略**: 保留未加密的备份文档
4. **审计跟踪**: 记录所有安全操作

## 相关功能

- [文档属性](./02-basic-operations.md) - 查看和设置文档属性
- [转换](./11-conversion.md) - 受保护文档的转换
