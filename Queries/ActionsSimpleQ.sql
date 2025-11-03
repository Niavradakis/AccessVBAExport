SELECT ActionsT.*, ActionTypesT.Action_Type_Description, ProtocolsSimpleQ.Protocol_Type_Description, ProtocolsSimpleQ.ProtocolsT.Priority_Level, 8 AS EntitiesTypesToHaveAttributes
FROM ProtocolsSimpleQ INNER JOIN (ActionsT INNER JOIN ActionTypesT ON ActionsT.Action_Type_ID = ActionTypesT.Action_Type_ID) ON ProtocolsSimpleQ.Protocol_ID = ActionsT.Protocol_ID;

