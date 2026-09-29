window.commandDefinition = [
    {
        commandName: "Get-Mailbox",
        tag: ["Mailbox", "MailboxMove"],
        parameters: [
            { key: "-Archive", type: "switch" },
            { key: "-SoftDeletedMailbox", type: "switch-split" },
            { key: "-IncludeInactiveMailbox", type: "switch" },
            { key: "-InactiveMailboxOnly", type: "switch-split" }   
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxStatistics",
        tag: ["Mailbox"],
        parameters: [
            { key: "-Archive", type: "switch-split" },
            { key: "-IncludeMoveHistory", type: "switch" },
            { key: "-IncludeMoveReport", type: "switch" }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxFolderStatistics",
        tag: ["Mailbox"],
        parameters: [
            { key: "-Archive", type: "switch-split" },
            { key: "-IncludeOldestAndNewestItems", type: "switch" },
            { key: "-IncludeAnalysis", type: "switch" },
            {
                key: "-FolderScope",
                type: "multi-checkbox-split",
                isMandatory: false,
                options: ["All","Archive","Calendar","Contacts","ConversationHistory","DeletedItems","Drafts","Inbox","JunkEmail","Journal","LegacyArchiveJournals","ManagedCustomFolder","NonIpmRoot","Notes","Outbox","Personal","RecoverableItems","RssSubscriptions","SentItems","SyncIssues","Tasks"]
            }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxFolderPermission",
        tag: ["Mailbox", "Calendar"],
        parameters: [
            {
                key: "Folder",
                type: "multi-checkbox-split",
                isMandatory: true,
                options: [
                    { value: "ルート", english: "Root", label: "ルート (Root)" },
                    { value: "予定表", english: "Calendar", label: "予定表 (Calendar)" },
                    { value: "受信トレイ", english: "Inbox", label: "受信トレイ (Inbox)" },
                    { value: "連絡先", english: "Contacts", label: "連絡先 (Contacts)" }
                ]
            }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxPermission",
        tag: ["Mailbox"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-RecipientPermission",
        tag: ["Mailbox", "Group", "Other recipient"],
        applicable: ["exo"]
    },
    {
        commandName: "Get-CASMailbox",
        tag: ["Mailbox"],
        parameters: [
            { key: "-ActiveSyncDebugLogging", type: "switch" }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Export-MailboxDiagnosticLogs",
        tag: ["Mailbox", "Calendar"],
        parameters: [
            {
                key: "-ComponentName",
                type: "multi-checkbox-split",
                isMandatory: true,
                options: ["AcceptCalendarSharingInvite", "AttendeeListReplicationAssistant", "AttendeesProperty", "BirthdayCalendar", "CalendarActiveSharees", "CalendarAssistant", "CalendarPermissions", "CalendarRepair", "CalendarReplication", "CalendarSharingEmailSync", "CalendarSharingInconsistencyRepair", "CalendarSharingInconsistencyValidator", "CalendarSharingInitialView", "CalendarSharingInvite", "CalendarSharingLocalFolder", "ClearCalendar", "DefaultViewIndexer", "DelegateRulesManagement", "FreeBusyPublishingAssistantQuickLog", "HoldTracking", "InternalCalendarSharingMigration", "InternetCalendar", "MeetingMessageProcessingAgent", "MFN", "MRM", "OnlineMeetings", "OOFRules", "RBA", "RemindersAssistant", "ReplicateAttendees", "SearchFoldersMigration", "SharedInTombstones", "Sharing", "SharingMigrationAssistant", "SharingSyncAssistant", "SmartTaggingFai", "SubstrateHoldTracking", "SweepRules", "TimeProfile", "TriggerSharingSyncAsNeeded", "VisibleMeetingMessageProcessing"]
            }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-User",
        tag: ["Mailbox", "Other recipient"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-Recipient",
        tag: ["Mailbox", "Group", "Other recipient"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailUser",
        tag: ["Other recipient", "MailboxMove"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailContact",
        tag: ["Other recipient"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-CalendarProcessing",
        tag: ["Mailbox", "Calendar"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-InboxRule",
        tag: ["Mailbox"],
        parameters: [
            { key: "-IncludeHidden", type: "switch" }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxMessageConfiguration",
        tag: ["Mailbox"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxRegionalConfiguration",
        tag: ["Mailbox"],
        parameters: [
            { key: "-VerifyDefaultFolderNameLanguage", type: "switch" }
        ],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MailboxJunkEmailConfiguration",
        tag: ["Mailbox"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-DistributionGroup",
        tag: ["Group"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-DistributionGroupMember",
        tag: ["Group"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-DynamicDistributionGroup",
        tag: ["Group"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-DynamicDistributionGroupMember",
        tag: ["Group"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-UnifiedGroup",
        tag: ["Group"],
        applicable: ["exo"]
    },
    {
        commandName: "Get-UnifiedGroupLinks",
        tag: ["Group"],
        parameters: [
            {
                key: "-LinkType",
                type: "multi-checkbox-split",
                isMandatory: true,
                options: ["Members", "Owners", "Subscribers"]
            }
        ],
        applicable: ["exo"]
    }
    ,
    {
        commandName: "Get-Group",
        tag: ["Group"],
        applicable: ["exo", "onprem"]
    },
    {
        commandName: "Get-MgUser",
        tag: ["Mailbox", "Other recipient"],
        parameters: [
            {
                key: "-Property",
                type: "multi-checkbox",
                isMandatory: false,
                defaultSelectedOptions: ["Id", "DisplayName", "UserPrincipalName", "Mail", "ServiceProvisioningErrors", "AssignedLicenses", "LicenseAssignmentStates", "OnPremisesLastSyncDateTime", "OnPremisesImmutableId"],
                options: ["AboutMe","AccountEnabled","Activities","AgeGroup","AgreementAcceptances","AppRoleAssignments","AssignedLicenses","AssignedPlans","Authentication","AuthorizationInfo","Birthday","BusinessPhones","Calendar","CalendarGroups","CalendarView","Calendars","Chats","City","CloudClipboard","CompanyName","ConsentProvidedForMinor","ContactFolders","Contacts","Country","CreatedDateTime","CreatedObjects","CreationType","CustomSecurityAttributes","DeletedDateTime","Department","DeviceManagementTroubleshootingEvents","DirectReports","DisplayName","Drive","Drives","EmployeeExperience","EmployeeHireDate","EmployeeId","EmployeeLeaveDateTime","EmployeeOrgData","EmployeeType","Events","Extensions","ExternalUserState","ExternalUserStateChangeDateTime","FaxNumber","FollowedSites","GivenName","HireDate","Id","Identities","ImAddresses","InferenceClassification","Insights","Interests","IsManagementRestricted","IsResourceAccount","JobTitle","JoinedTeams","LastPasswordChangeDateTime","LegalAgeGroupClassification","LicenseAssignmentStates","LicenseDetails","Mail","MailFolders","MailNickname","ManagedAppRegistrations","ManagedDevices","Manager","MemberOf","Messages","MobilePhone","MySite","Oauth2PermissionGrants","OfficeLocation","OnPremisesDistinguishedName","OnPremisesDomainName","OnPremisesExtensionAttributes","OnPremisesImmutableId","OnPremisesLastSyncDateTime","OnPremisesProvisioningErrors","OnPremisesSamAccountName","OnPremisesSecurityIdentifier","OnPremisesSyncEnabled","OnPremisesUserPrincipalName","Onenote","OnlineMeetings","OtherMails","Outlook","OwnedDevices","OwnedObjects","PasswordPolicies","PasswordProfile","PastProjects","People","PermissionGrants","Photo","Photos","Planner","PostalCode","PreferredDataLocation","PreferredLanguage","PreferredName","Presence","Print","ProvisionedPlans","ProxyAddresses","RegisteredDevices","Responsibilities","Schools","ScopedRoleMemberOf","SecurityIdentifier","ServiceProvisioningErrors","Settings","ShowInAddressList","SignInActivity","SignInSessionsValidFromDateTime","Skills","Solutions","Sponsors","State","StreetAddress","Surname","Teamwork","Todo","TransitiveMemberOf","UsageLocation","UserPrincipalName","UserType","AdditionalProperties"]
            }
        ],
        applicable: ["graph"]
    },
    {
        commandName: "Get-MgUserLicenseDetail",
        tag: ["Mailbox", "Other recipient"],
        applicable: ["graph"]
    },
    {
        commandName: "Get-RemoteMailbox",
        tag: ["Mailbox", "Other recipient", "MailboxMove"],
        applicable: ["onprem"]
    },
    {
        commandName: "Get-ADUser",
        tag: ["Mailbox", "Other recipient"],
        applicable: ["onprem"]
    },
    {
        commandName: "Get-ADGroup",
        tag: ["Group"],
        applicable: ["onprem"]
    },
    {
        commandName: "Get-ADObject",
        tag: ["Mailbox", "Group", "Other recipient"],
        applicable: ["onprem"]
    },
    {
        commandName: "Get-ADPermission",
        tag: ["Mailbox", "Group", "Other recipient"],
        applicable: ["onprem"]
    },
    {
        commandName: "Get-OrganizationConfig",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-TransportConfig",
        tag: ["Organization", "Transport"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-TransportRule",
        tag: ["Organization", "Transport"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-InboundConnector",
        tag: ["Organization", "Transport"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-OutboundConnector",
        tag: ["Organization", "Transport"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-AcceptedDomain",
        tag: ["Organization", "Transport"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-RemoteDomain",
        tag: ["Organization", "Transport"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-OrganizationRelationship",
        tag: ["Organization", "Calendar"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-SharingPolicy",
        tag: ["Organization", "Calendar"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-IntraOrganizationConnector",
        tag: ["Organization", "Calendar"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-MigrationBatch",
        tag: ["Organization", "MailboxMove"],
        applicable: ["exo", "onprem"],
        orgLevel: true,
        parameters: [
            { key: "-IncludeReport", type: "switch" }
        ]
    },
    {
        commandName: "Get-MoveRequest",
        tag: ["MailboxMove"],
        applicable: ["exo", "onprem"],
        parameters: []
    },
    {
        commandName: "Get-MoveRequestStatistics",
        tag: ["MailboxMove"],
        applicable: ["exo", "onprem"],
        parameters: [
            { key: "-IncludeReport", type: "switch" }
        ]
    },
    {
        commandName: "Get-MigrationUser",
        tag: ["MailboxMove"],
        applicable: ["exo", "onprem"],
        parameters: []
    },
    {
        commandName: "Get-MigrationUserStatistics",
        tag: ["MailboxMove"],
        applicable: ["exo", "onprem"],
        parameters: [
            { key: "-IncludeReport", type: "switch" }
        ]
    },
    {
        commandName: "Get-RetentionPolicy",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-RetentionPolicyTag",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        parameters: [
            { key: "-IncludeSystemTags", type: "switch" }
        ],
        orgLevel: true
    },
    {
        commandName: "Get-AddressBookPolicy",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-GlobalAddressList",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-OfflineAddressBook",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    },
    {
        commandName: "Get-AddressList",
        tag: ["Organization"],
        applicable: ["exo", "onprem"],
        orgLevel: true
    }
];
