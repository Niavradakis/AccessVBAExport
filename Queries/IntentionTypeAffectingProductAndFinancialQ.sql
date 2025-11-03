SELECT IntentionsTypeT.Intention_Type_Description, IntentionsTypeT.Intention_Type_ID
FROM IntentionsTypeT
WHERE (((IntentionsTypeT.Involves_Products)=Yes) AND ((IntentionsTypeT.Involves_Financial)=Yes))
ORDER BY IntentionsTypeT.View_ID;

