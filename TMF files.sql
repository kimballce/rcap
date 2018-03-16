SELECT a.Id, a.[Name] AS [FileName], a.FileType, a.Note, a.[Date], a.Category, a.[By], pxt.ProjectId, pxt.TaskStatus, p.[Name] AS ProjectName, p.ZipCodes, p.[Status], f.[Name] AS [FundingProgram], fe.[Name] AS [FundingEntity]
FROM Attachment a
JOIN Project_X_Task pxt ON a.Id = pxt.AttachmentId
JOIN Project p ON pxt.ProjectId = p.Id
JOIN Funding f ON p.FundingProgId = f.Id
JOIN FundingEntity fe ON f.FunderEntityId = fe.Id
WHERE a.Category = 'TMF Assessment'