SELECT ProtocolsT.*, ProtocolTypesT.Protocol_Type_Description, ProtocolsT.Priority_Level, 9 AS EntityTypesToHaveAttributesID
FROM ProtocolsT INNER JOIN ProtocolTypesT ON ProtocolsT.Protocol_Type_ID = ProtocolTypesT.Protocol_Type_ID;

