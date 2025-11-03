SELECT ProtocolsT.*, ActionsT.*, AttributesAndSubattributesForProtocolsQ.*, AttributesAndSubAttributesForActionsQ.*
FROM (ProtocolsT LEFT JOIN AttributesAndSubattributesForProtocolsQ ON ProtocolsT.Protocol_ID = AttributesAndSubattributesForProtocolsQ.Entity_ID) INNER JOIN (ActionsT LEFT JOIN AttributesAndSubAttributesForActionsQ ON ActionsT.Action_ID = AttributesAndSubAttributesForActionsQ.Entity_ID) ON ProtocolsT.Protocol_ID = ActionsT.Protocol_ID;

