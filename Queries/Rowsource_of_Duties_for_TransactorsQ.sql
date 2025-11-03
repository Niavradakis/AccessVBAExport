SELECT DutiesT.Duty_ID, DutiesT.Duty_Description, TransactorsT.Transactor_ID
FROM TransactorsT INNER JOIN ((JobsT INNER JOIN (DutiesT INNER JOIN DutiesPerJobT ON DutiesT.Duty_ID = DutiesPerJobT.Duty_ID) ON JobsT.Job_ID = DutiesPerJobT.Job_ID) INNER JOIN LinkTransactorsToJobsT ON JobsT.Job_ID = LinkTransactorsToJobsT.Job_ID) ON TransactorsT.Transactor_ID = LinkTransactorsToJobsT.Transactor_ID
WHERE (((TransactorsT.Transactor_ID)=[forms].[LinkTransactorToDutiesF].[Transactor_IDCbo]));

