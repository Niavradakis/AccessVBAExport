SELECT CustomersFromMitsosDatabaseT.ref_transactor_id, CustomersFromMitsosDatabaseT.customer_email, CustomersFromMitsosDatabaseT.customer_telephone, CustomersFromMitsosDatabaseT.customer_mobile, IIf([customer_address] & ""="",[customer_city],[customer_address] & ", " & [customer_city]) AS Full_Address, CustomersFromMitsosDatabaseT.customer_vat_number, CustomersFromMitsosDatabaseT.customer_company_name, CustomersFromMitsosDatabaseT.customer_irs, CustomersFromMitsosDatabaseT.customer_postal_code
FROM CustomersFromMitsosDatabaseT
WHERE (((CustomersFromMitsosDatabaseT.customer_country_id)=1));

