select count(*) as TotalRapidOrders from NETSERVER3.RTMGMT.dbo.ORDERS
select count(*) as TotalUnisonOrders from UN_ORDERS.dbo.ORDERS
select count(*) as UnprocessedOrderQueueRows from NETSERVER3.RTMGMT.dbo.ORDERS_Q WHERE PROCESSED = 'N'
select count(*) as TotalRapidCODs from NETSERVER3.RTMGMT.dbo.ORDERCOD
select count(*) as TotalUnisonCODs from UN_ORDERS.dbo.ORDERCOD
select count(*) as UnprocessedCODQueueRows from NETSERVER3.RTMGMT.dbo.ORDERCOD_Q WHERE PROCESSED = 'N'

select count(*) as UnisonCustomers from unison.dbo.customer --where status = 1

select count(*) as RapidCustomers from un_orders.dbo.customer --where CreditStatus <> 'F'

select distinct(creditstatus) from un_orders.dbo.customer

-- Accounts that are in Rapid, but not in Unison
select * from un_orders.dbo.customer where [id] not in (select [id] from unison.dbo.customer) --and creditstatus <> 'F'

-- Accounts that are in Unison, but not in Rapid
select * from unison.dbo.customer where [id] not in (select [id] from un_orders.dbo.customer) and status = 1

-- Accounts that are in both Unison & Rapid
select count(*) from un_orders.dbo.customer where creditstatus <> 'F'
select * from un_orders.dbo.customer where [id] in (select [id] from unison.dbo.customer where status = 1) and creditstatus <> 'F'
select * from un_orders.dbo.customer where [id] in (select [id] from unison.dbo.customer where status <> 1) and creditstatus <> 'F'
select * from un_orders.dbo.customer where [id] not in (select [id] from unison.dbo.customer) and creditstatus <> 'F'
select c.[id] as CustomerID, a.[name] as CustomerName, a.CONTACT from un_orders.dbo.Customer c, un_orders.dbo.Address a where (c.[id] = a.ownerid) and (c.addressid = a.[id]) and (c.creditstatus <> 'F') and (a.contact not like '%CLOSE%') and c.[id] not in (select [id] from unison.dbo.customer)

select c.[id] as CustomerID, a.[name] as CustomerName, a.CONTACT from un_orders.dbo.Customer c, un_orders.dbo.Address a where (c.[id] = a.ownerid) and (c.addressid = a.[id]) and (c.creditstatus <> 'F') and (a.contact not like '%CLOSE%') order by CustomerID
select c.[id] from un_orders.dbo.Customer c, un_orders.dbo.Address a where (c.[id] = a.ownerid) and (c.addressid = a.[id]) and (c.creditstatus <> 'F') and (a.contact not like '%CLOSE%') and (a.name not like '*%') order by [ID]


select * from unison.dbo.customer where [id] in (select [id] from un_orders.dbo.customer) and status = 1


select * from unison.dbo.customer where [id] in (select c.[id] from un_orders.dbo.Customer c, un_orders.dbo.Address a where (c.[id] = a.ownerid) and (c.addressid = a.[id]) and (c.creditstatus <> 'F') and (a.contact not like '%CLOSE%') and (a.name not like '*%')) and status = 0 order by [id]
select * from unison.dbo.customer where [id] in (select c.[id] from un_orders.dbo.Customer c, un_orders.dbo.Address a where (c.[id] = a.ownerid) and (c.addressid = a.[id]) and (c.creditstatus <> 'F') and (a.contact not like '%CLOSE%') and (a.name not like '*%')) and status = 1
