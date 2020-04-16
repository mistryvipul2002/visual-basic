


create table  incometax (
      emp_id number(10)  unique not null,
      emp_first_name varchar2(15) ,
      emp_last_name varchar2(15) ,
      emp_insti varchar2(15),
      emp_dept  varchar2(15),
      emp_basic_pay number(8) ,
      emp_ha number(8),
      emp_ta number(8), 
      emp_other_allow number(8),
      emp_rebates number(8)
        );
create table  var (
      emp_id number(10)  unique not null,     
      emp_da number(8), 
      emp_gross number(8),
      emp_income_tax number(8)
        );

insert into incometax ( emp_id ,emp_first_name,emp_last_name,
emp_insti,emp_dept,emp_basic_pay,emp_ha,emp_ta,emp_other_allow,emp_rebates)
values(730,'sats','kunta','ism','cse',8000,450,300,200,350);

insert into incometax ( emp_id ,emp_first_name,emp_last_name,
emp_insti,emp_dept,emp_basic_pay,emp_ha,emp_ta,emp_other_allow,emp_rebates)
values(7770,'sats','kunta','ism','cse',8000,450,300,200,350);

.....................to create  trigger..............
create or replace trigger trig1 
   after insert or update or delete on incometax for each row
begin
      if inserting then
                  if :new.emp_basic_pay <10000 then
                       	insert into var( emp_id,emp_da,emp_gross,emp_income_tax) values(:new.emp_id ,:new.emp_basic_pay*1.4,:new.emp_basic_pay*12+:new.emp_basic_pay*1.4*12+:new.emp_ha*12+:new.emp_ta+:new.emp_other_allow-:new.emp_rebates,0 ) ;
                  elsif :new.emp_basic_pay <20000 then                      
                   	insert into var( emp_id,emp_da,emp_gross,emp_income_tax) values(:new.emp_id ,:new.emp_basic_pay*1.2,:new.emp_basic_pay*12+:new.emp_basic_pay*1.2*12+:new.emp_ha*12+:new.emp_ta+:new.emp_other_allow-:new.emp_rebates,0  ) ; 
                  elsif :new.emp_basic_pay <30000 then
			insert into var( emp_id,emp_da,emp_gross,emp_income_tax) values(:new.emp_id ,:new.emp_basic_pay*.8,:new.emp_basic_pay*12+:new.emp_basic_pay*.8*12+:new.emp_ha*12+:new.emp_ta+:new.emp_other_allow-:new.emp_rebates,0  ) ;
                  else 
			insert into var( emp_id,emp_da,emp_gross,emp_income_tax) values(:new.emp_id ,:new.emp_basic_pay*.2,:new.emp_basic_pay*12+:new.emp_basic_pay*.2*12+:new.emp_ha*12+:new.emp_ta+:new.emp_other_allow-:new.emp_rebates,0  ) ;
                  end if;
                  update var set emp_income_tax = (emp_gross-50000)*.2 where
			emp_id=:new.emp_id and emp_gross<60000 and emp_gross>50000;
                  update var set emp_income_tax = (emp_gross-60000)*.3 where
			emp_id=:new.emp_id and emp_gross<150000 and emp_gross>60000;
                  update var set emp_income_tax = (emp_gross-150000)*.5 where
			emp_id=:new.emp_id and emp_gross>150000;
      elsif deleting then
		   delete from var where emp_id=:old.emp_id;
      end if;
end trig1;
.............................................................
create or replace view mastertable (emp_id,emp_first_name,emp_last_name,emp_dept,emp_basic_pay,emp_gross,emp_income_tax )as (
     select incometax.emp_id,incometax.emp_first_name,incometax.emp_last_name,incometax.emp_dept,incometax.emp_basic_pay,var.emp_gross,var.emp_income_tax 
from incometax,var
where incometax.emp_id=var.emp_id);
..............................................................