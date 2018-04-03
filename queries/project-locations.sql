SELECT p.Id, p.[Name], p.[Status], p.ZipCodes, LEFT(p.ZipCodes, 5) AS [FirstZip], p.HasNoPlaceZips, p.[State], f.[Name] AS [ProgramName], p.AssistantType1, p.StartDate, p.EndDate, fe.[Name] AS [FundingEntity]
FROM Project p
JOIN Funding f ON f.Id = p.FundingProgId
JOIN FundingEntity fe ON fe.Id = f.FunderEntityId
WHERE p.[Status] <> 'Cancelled'
AND fe.[Name] IN ('HHS/OCS', 'US EPA', 'USDA Rural Development')
