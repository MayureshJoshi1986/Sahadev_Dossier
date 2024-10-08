USE [ADF_SAHADEV]
GO
SET IDENTITY_INSERT [dbo].[mstCoverageDossierStatus] ON 

INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (1, N'Pending', CAST(N'2024-03-18T15:43:00' AS SmallDateTime), NULL, N'Dossier Request Added')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (2, N'Review Pending', CAST(N'2024-03-18T15:43:00' AS SmallDateTime), NULL, N'Enrichment will Completed(Data Count will match with Enrichment table or data will collect) then all article will insert in new table with deleted False status , then status will to Review Pending')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (3, N'Review Completed', CAST(N'2024-03-18T15:43:00' AS SmallDateTime), NULL, N'User Will select particular Article from Dossier UI & approve, then Status will change from Review Pending to Review Completed')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (4, N'Summary Started', CAST(N'2024-03-18T15:43:00' AS SmallDateTime), NULL, N'System will start creating summary of article which is selected from user then status will change from Review Completed To Summary Started')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (5, N'Summary Completed', CAST(N'2024-03-18T15:43:00' AS SmallDateTime), NULL, N'System will generate summary of all selected article then status will change Summary Started To Summary Completed')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (6, N'Dossier Generation Start', CAST(N'2024-03-26T15:12:00' AS SmallDateTime), NULL, N'When status will change summary completed then Dossier Generation Start. Status will change From Summary Completed to Dossier Generation Start')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (7, N'Dossier Generation Completed', CAST(N'2024-03-26T15:12:00' AS SmallDateTime), NULL, N'When Dossier file will generation completed then status will change from Dossier Generation start to Competed')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (8, N'Email Sent', CAST(N'2024-03-26T15:12:00' AS SmallDateTime), NULL, N'File will generate &Email will sent to the user then Status will change Dossier Generation Completed to Email Sent')
INSERT [dbo].[mstCoverageDossierStatus] ([NLStatusID], [Status], [CreatedAt], [ModifiedAt], [Description]) VALUES (9, N'Additional Url', CAST(N'2024-04-24T16:47:00' AS SmallDateTime), NULL, NULL)
SET IDENTITY_INSERT [dbo].[mstCoverageDossierStatus] OFF
GO
