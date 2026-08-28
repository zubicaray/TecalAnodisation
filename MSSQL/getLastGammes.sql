select  NumGammeAnodisation,count(*) from  DetailsChargesProduction 
where 	DateEntreeEnLigne >=  '20260101'  
group by NumGammeAnodisation
order by count(*) desc