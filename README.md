# Arcane Server

![Banner](assets/imgs/banner.png)

This repository contains the Arcane Server component of the Arcane project, which is fully implemented in PowerShell. It operates independently without relying on any third-party software such as RDP or VNC. Instead, it leverages the native Windows API, using the full capabilities of PowerShell.

> ⓘ Since version 1.0.4, the Arcane Viewer and Server have separate versioning, allowing each to progress independently. This separation ensures that if only Viewer features are optimized, enhanced, or bug-fixed, the Server version doesn't need to be updated unnecessarily (and vis-versa). Although having different versions for the Viewer and Server might seem confusing, the key detail to focus on is the protocol version. The protocol version determines compatibility between the Viewer and Server, ensuring they work together correctly.

## Quick Setup (Latest Release) - [PowerShell Gallery](https://www.powershellgallery.com)

> ⚠️ Please note that you must have administrative privileges to install a new PowerShell module.

Open an elevated PowerShell prompt and execute the following command:

```powershell
Install-Module -Name Arcane_Server
```

The latest version of the Arcane Server should now be installed and available.

Before running the server, you must import the module into your current PowerShell session, note that it is now mandatory to have an elevated PowerShell session, Arcane Server support both running as limited and privileged user, however, if session is running with limited privilege, mouse and keyboard wont be able to be captured for elevated window's

> ⓘ depending on your system configuration, you may need to run the following command to temporarily bypass the execution policy in order to run an unsigned script:
> `powershell.exe -executionpolicy bypass`

```powershell
Import-Module Arcane_Server
```

Once the module is imported, you can run the server using the following command:

```powershell
Invoke-ArcaneServer
```

That's it, you're ready to go! 🚀

## Capture LogonUI / UAC (Secure Desktop)

Starting with version `1.0.5` of Arcane Server, **Secure Desktop** is fully supported using just a single instance of the server. This enhancement allows you to log in to your computer directly from Arcane or respond to UAC (User Account Control) prompts. This feature is crucial for those who wish to use Arcane as a day-to-day remote desktop application.

In the near future, I will publish an article detailing how I implemented this feature without relying on third-party services, unlike other remote desktop applications.

To support **Secure Desktop** capture, the Server must be run as an Interactive **NT/Authority SYSTEM** process. "Interactive" means a SYSTEM process that has access to the active desktop session you wish to capture. Tools like **PsExec** can facilitate this by spawning a separate interactive process as SYSTEM. However, PsExec can sometimes be flagged as malicious, as it's frequently used by threat actors and red teamers.

Fortunately, a few years ago, I developed a PowerShell script called [PowerRunAsSystem](https://github.com/PhrozenIO/PowerRunAsSystem). This script allows you to spawn an interactive SYSTEM process using only native Windows functions, without relying on external tools. You can install **PowerRunAsSystem** directly via the PowerShell Gallery:

> ⚠️ Please note that you must have administrative privileges to install a new PowerShell module.

```powershell
Install-Module -Name PowerRunAsSystem
```

In the same PowerShell session or a new one with administrative privileges, import the newly installed module using:

> ⓘ depending on your system configuration, you may need to run the following command to temporarily bypass the execution policy in order to run an unsigned script:
> `powershell.exe -executionpolicy bypass`

```powershell
Import-Module PowerRunAsSystem
```

Now you can call:

```powershell
 Invoke-InteractiveSystemPowerShell
```

A new PowerShell command prompt should open with SYSTEM privileges. You can verify this by running the command `whoami`. From this prompt, you can now start your Arcane Server as you would in a regular prompt. When Arcane Server is run under the SYSTEM user account, it automatically detects this and enables Secure Desktop interaction capabilities.

## Version Table

| Version         | Protocol Version | Release Date    | 
|-----------------|------------------|-----------------|
| 1.0.4           | 5.0.1            | 15 August 2024  | 

## Advanced Usage

```powershell
Invoke-ArcaneServer
```

### Supported Options:
 
| Parameter              | Type             | Default    | Description  |
|------------------------|------------------|------------|--------------|
| ServerAddress          | String           | 0.0.0.0    | IP address representing the local machine's IP address |
| ServerPort             | Integer          | 2801       | The port number on which to listen for incoming connections |
| SecurePassword         | SecureString     | None       | SecureString object containing the password used for authenticating remote viewers (recommended) |
| Password               | String           | None       | Plain-text password used for authenticating remote viewers (not recommended; use SecurePassword instead) |
| DisableVerbosity       | Switch           | False      | If specified, the program will suppress verbosity messages |
| UseTLSv1_3             | Switch           | False      | If specified, the program will use TLS v1.3 instead of TLS v1.2 for encryption (recommended if both systems support it) |
| Clipboard              | Enum             | Both       | Specify the clipboard synchronization mode (options include 'Both', 'Disabled', 'Send', and 'Receive'; see below for more detail) |
| CertificateFile        | String           | None       | A file containing valid certificate information (x509) that includes the private key  |
| EncodedCertificate     | String           | None       | A base64-encoded representation of the entire certificate file, including the private key |
| ViewOnly               | Switch           | False      | If specified, the remote viewer will only be able to view the desktop and will not have access to the mouse or keyboard |
| PreventComputerToSleep | Switch           | False      | If specified, this option will prevent the computer from entering sleep mode while the server is active and waiting for new connections |
| CertificatePassword    | SecureString     | None       | Specify the password used to access a password-protected x509 certificate provided by the user | 

### Server Address Examples

| Value             | Description                                                              | 
|-------------------|--------------------------------------------------------------------------|
| 127.0.0.1         | Only listen for connections from the localhost (usually for debugging purposes) |
| 0.0.0.0           | Listen for connections on all network interfaces, including the local network and the internet                       |

### Clipboard Mode Enum Properties

| Value             | Description                                        | 
|-------------------|----------------------------------------------------|
| Disabled          | Clipboard synchronization is disabled on both the viewer and server sides |
| Receive           | Only incoming clipboard data is allowed                |
| Send              | Only outgoing clipboard data is allowed                 |
| Both              | Clipboard synchronization is allowed on both the viewer and server sides  |

### ⚠️ Important Notices

1. It is recommended to use SecurePassword instead of a plain-text password, even if the plain-text password is being converted to a SecureString.
2. If you do not specify a custom certificate using 'CertificateFile' or 'EncodedCertificate', a default self-signed certificate will be generated and installed for the local user.
3. If you do not specify a SecurePassword or Password, a random, complex password will be generated and displayed in the terminal (this password is temporary).

### Examples

```powershell
Invoke-ArcaneServer -ListenAddress "0.0.0.0" -ListenPort 2801 -SecurePassword (ConvertTo-SecureString -String "urCompl3xP@ssw0rd" -AsPlainText -Force)

Invoke-ArcaneServer -ListenAddress "0.0.0.0" -ListenPort 2801 -SecurePassword (ConvertTo-SecureString -String "urCompl3xP@ssw0rd" -AsPlainText -Force) -CertificateFile "c:\certs\phrozen.p12"
```

### Generate your Certificate

```
openssl req -x509 -sha512 -nodes -days 365 -newkey rsa:4096 -keyout phrozen.key -out phrozen.crt
```

Then export the new certificate (**must include private key**).

```
openssl pkcs12 -export -out phrozen.p12 -inkey phrozen.key -in phrozen.crt
```

### Integrate to server as a file

Use `CertificateFile`. Example: `c:\tlscert\phrozen.crt`

### Integrate to server as a base64 representation

Encode an existing certificate using PowerShell

```powershell
[convert]::ToBase64String((Get-Content -path "c:\tlscert\phrozen.crt" -Encoding byte))
```
or on Linux / Mac systems

```
base64 -i /tmp/phrozen.p12
```

You can then pass the output base64 certificate file to parameter `EncodedCertificate` (One line)

## Changelog


