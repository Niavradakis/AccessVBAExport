SELECT DutiesT.Duty_ID, DutiesT.Duty_Description, JobsT.Job_Title, LinkTransactorsToJobsT.Transactor_ID
FROM (JobsT INNER JOIN (DutiesT INNER JOIN LInkDutiesPerJobT ON DutiesT.Duty_ID = LInkDutiesPerJobT.Duty_ID) ON JobsT.Job_ID = LInkDutiesPerJobT.Job_ID) INNER JOIN LinkTransactorsToJobsT ON JobsT.Job_ID = LinkTransactorsToJobsT.Job_ID
WHERE (((LinkTransactorsToJobsT.Transactor_ID)=[Forms]![LinkTransactorToDutiesF]![Transactor_IDCbo]));

