```flow
st=>start: Start
op_excel=>operation: Choose Excel File
cond_ex=>condition: yes or no?

op_test=>operation: Send Test E-mails?
cond_test=>condition: yes or no?

op_send_test=>operation: Send Test E-mails

op_real=>operation: Send E-mails for Real?
cond_real=>condition: yes or no?
send_real=>operation: Send E-mails for Real.
e=>end

st->op_excel->cond_ex
cond_ex(no)->e
cond_ex(yes)->op_test->cond_test
cond_test(yes)->op_send_test->e
cond_test(no)->op_real->cond_real
cond_real(no)->e
cond_real(yes)->send_real
```