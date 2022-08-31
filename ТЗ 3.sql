with g as
(
select 
		id, 
		data, reason, 
		RANK() OVER (PARTITION BY id,reason ORDER BY YEAR(h.DATA), MONTH(DATA), DAY(DATA)) AS flag
from h
)
select 
		t.id, 
		avg(DATEDIFF(minute, t.data ,t2.data)) as �������  
from g t
join g t2 on t2.id = t.id and t2.reason = '��������' and t.flag=t2.flag
where t.reason = '�������'
group by t.id

----------------------------------------------------------------------------------------------------------------------

if object_id('tempdb..#t') is not null drop table #t
select 
		id, 
		data, reason, 
		RANK() OVER (PARTITION BY id,reason ORDER BY YEAR(h.DATA), MONTH(DATA), DAY(DATA)) AS flag
into #t
from h
order by data, flag

select 
		t.id, 
		avg(DATEDIFF(minute, t.data ,t2.data)) as �������  
from #t t
join #t t2 on t2.id = t.id and t2.reason = '��������' and t.flag=t2.flag
where t.reason = '�������'
group by t.id

----------------------------------------------------------------------------------------------------------------------

select 
		t.id, 
		avg(DATEDIFF(minute, t.data ,t2.data)) as �������  
from h t
outer apply (select top 1 * from h t2 where t2.id = t.id and t2.reason = '��������' and t.data<t2.data) as t2
where t.reason = '�������'
group by t.id