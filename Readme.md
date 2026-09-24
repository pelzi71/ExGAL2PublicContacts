# ExGAL2PublicContacts

Synchronize Exchange Global Address List users with Exchange Public Folder Contacts.

## About

ExGAL2PublicContacts is a PowerShell-based tool that synchronizes users from an Exchange Global Address List (GAL) with contacts stored in an Exchange Public Folder.

The project was created to solve a practical Exchange administration requirement that is not directly provided as a built-in Exchange feature.

## Current Status

**Development version 0.1.0**

The current version is **read-only**.

It reads Exchange mailboxes, Active Directory account status and existing Public Folder contacts, but does not create, modify or delete contacts.

## Target Environment

- Exchange Server 2019
- Exchange Server 2019 SE
- PowerShell 5.1
- EWS Managed API
- Active Directory

## Main Concept

```text
Exchange / Active Directory
           |
           | UserMailbox
           | AD account status
           v
       ExGAL2PublicContacts
           |
           | SMTP address matching
           v
Exchange Public Folder
       Contacts