job_years = 3 
salary = 200000
message = 'Addition is equal {coef}. New salary is {new_salary}!'
if job_years < 1:
    coefficient = 0
elif job_years >= 1 and job_years <= 2:
    coefficient = 0.2
else:
    coefficient = 0.4
new_salary = salary + salary * coefficient
print(message.format(coef=coefficient,new_salary=new_salary))