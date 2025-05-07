create elon identified by musk;
grant resource, connect to elon;
grant dba to elon;
conn elon/musk;

CREATE TABLE customer (
    cus_id VARCHAR2(10),
    cus_name VARCHAR2(20),
    addr VARCHAR2(20),
    phone_no NUMBER(12),
    cus_bal NUMBER(9,2)
);

CREATE TABLE product (
    pro_id VARCHAR2(10),
    pro_name VARCHAR2(25),
    pro_type VARCHAR2(25),
    mfg_rate NUMBER(7,2),
    pro_sell_rate NUMBER(7,2),
    stock_qty NUMBER(7)
);

CREATE TABLE raw_material (
    raw_id VARCHAR2(12),
    name VARCHAR2(20),
    company_name VARCHAR2(20),
    type VARCHAR2(20),
    unit VARCHAR2(20),
    size NUMBER(5),
    rate NUMBER(10),
    stock_qty NUMBER(7)
);

CREATE TABLE pro_master (
    order_no VARCHAR2(25),
    order_date VARCHAR2(20),
    supp_id VARCHAR2(10),
    method_of_payment VARCHAR2(10),
    advance NUMBER(15,2),
    dues NUMBER(15,2),
    net_amt NUMBER(15,2)
);

CREATE TABLE sales (
    inv_no VARCHAR2(10),
    cus_id VARCHAR2(10),
    sale_date VARCHAR2(20),
    method_of_payment VARCHAR2(10),
    adv NUMBER(15,6),
    net_amt NUMBER(15,6),
    total_qty NUMBER(5),
    dues NUMBER(15,6)
);

CREATE TABLE sales_details (
    inv_no VARCHAR2(10),
    pro_id VARCHAR2(10),
    pro_name VARCHAR2(50),
    qty NUMBER(5),
    rate NUMBER(15,6),
    gst NUMBER(10,8),
    amt NUMBER(15,6)
);

CREATE TABLE sup_det (
    supp_id VARCHAR2(20),
    name VARCHAR2(20),
    addr VARCHAR2(20),
    cont_no NUMBER(12),
    comp_name VARCHAR2(20)
);
