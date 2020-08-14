# JAVA
//标准导出
 @RequestMapping("download-transport")
    public void downloadTransport(HttpServletResponse response) {
        try {
            List<ExcelData> dataList = new ArrayList<>();
            ExcelData data = new ExcelData();
            List<List<String>> titles = new ArrayList<>();
            List<String> title = new ArrayList<>();


            data.setName("批量导入运单模板");
            
            List heads=new ArrayList();
            List<ExcelFontStyle> fontStyleList=new ArrayList<>();
            List<String> head=new ArrayList<>();
            head.add("教学诊改数据清单");
            ExcelFontStyle style=new ExcelFontStyle();
            style.setFontHeightInPoints((short) 14);
            style.setColor(IndexedColors.LIGHT_BLUE.getIndex()); 
            style.setFillForegroundColor(new Color(255,255,255));
            fontStyleList.add(style);
            heads.add(head)
            
             List<ExcelMergedRegion> excelMergedRegionList=new ArrayList<>();
            
            ExcelMergedRegion excelMergedRegion=new ExcelMergedRegion();
            excelMergedRegion.setFirstRow(0);
            excelMergedRegion.setLastRow(0);
            excelMergedRegion.setFirstCol(0);
            excelMergedRegion.setLastCol(12);
            excelMergedRegionList.add(excelMergedRegion);
            data.setHeadMergedRegions(excelMergedRegionList);
            
            
            title.add("订单编号");
            title.add("物流公司名称");
            title.add("物流公司编码");
            title.add("快递单号");
            //加入多行头部
            titles.add(title);
            data.setTitles(titles);
            //list
            List<List<Object>> rows = new Vector<>();
            List<Object> row = new ArrayList<>();  //行内有合并行列 请用List<ExclMultilieRows>
            row.add("12345678");
            row.add("中通速递");
            row.add("zhongtong");
            row.add("78564311");
            rows.add(row);
            //设置每列宽带
            Integer widths[]=new Integer[]{3300, 5300,0,0};
            List<Integer> colwidth = new ArrayList<>();
            colwidth.addAll(colWidths.apply(widths));
            data.setColsWidth(colwidth);
            //设置行高
            data.setColsHight(Short.parseShort("450"));

            data.setRows(rows);
            dataList.add(data);
            ExcelUtils.exportExcel(response, "批量导入运单模板.xlsx", dataList);
        } catch (Exception e) {
            logger.error("导出订单列表错误！", e);


        }
    }
    
    
   //## 报表统计
   /**
     * 客户统计
     *
     * @return
     */
    @ResponseBody
    @RequestMapping(value = "customer-statistics-json")
    public GridJsonResult jsonCustomerStatistics() {
        Page webPager = new Page();

        // 线程数量
        final int thread_num = 4;
        //原子性，保证一致性
        AtomicInteger atomicInteger = new AtomicInteger(0);
        Instant inst1 = Instant.now();  //当前的时间


        //建立3个线程
        CountDownLatch countDownLatch = new CountDownLatch(thread_num);
        ExecutorService executorService = Executors.newFixedThreadPool(thread_num);


        //  member_count;                   //会员总数
        //  order_member_count;         //下单会员总数
        //  order_count;                        //订单总数
        //  order_amount;                //订单总额

        // pay_ratio;                        //会员购买率
        // average_count;               //每会员订单数
        // average_amount;           //每会员购物额
        CustomerStatisticsDTO customerStatisticsDTO = new CustomerStatisticsDTO();


        //统计会员数量
        CompletableFuture.runAsync(() -> {
            CountAmountDTO countAmountDTO = reportManager.memberCount();
            customerStatisticsDTO.setMember_count(countAmountDTO.getCount());
            atomicInteger.incrementAndGet();
            countDownLatch.countDown();
        }, executorService);


        //统计下过订单的会员总数
        CompletableFuture.runAsync(() -> {
            CountAmountDTO countAmountDTO = reportManager.memberOrderCount();
            customerStatisticsDTO.setMember_order_count(countAmountDTO.getCount());
            atomicInteger.incrementAndGet();
            countDownLatch.countDown();
        }, executorService);


        //统计会员订单总数
        CompletableFuture.runAsync(() -> {
            CountAmountDTO countAmountDTO = reportManager.orderMemberCount();
            customerStatisticsDTO.setOrder_member_count(countAmountDTO.getCount());
            atomicInteger.incrementAndGet();
            countDownLatch.countDown();
        }, executorService);


        //统计订单数量和金额
        CompletableFuture.runAsync(() -> {
            CountAmountDTO countAmountDTO = reportManager.orderCountAmount();
            customerStatisticsDTO.setOrder_count(countAmountDTO.getCount());
            customerStatisticsDTO.setOrder_amount(countAmountDTO.getAmount());
            atomicInteger.incrementAndGet();
            countDownLatch.countDown();
        }, executorService);



        try

        {
            countDownLatch.await();
//            while (atomicInteger.get() < thread_num) {
//                if (Duration.between(inst1, Instant.now()).getSeconds() > 100) {
//                    break;
//                }
//                Thread.sleep(1000);
//
//
//            }


        } catch (
                Exception e)

        {

            logger.error("等待异常！！！");

        }


        customerStatisticsDTO.setPay_ratio( (new BigDecimal(customerStatisticsDTO.getMember_order_count()).divide(new BigDecimal(customerStatisticsDTO.getMember_count()),6,RoundingMode.HALF_UP)).doubleValue());
        customerStatisticsDTO.setAverage_count( (new BigDecimal(customerStatisticsDTO.getOrder_member_count()).divide(new BigDecimal(customerStatisticsDTO.getMember_count()),6,RoundingMode.HALF_UP)).doubleValue());
        customerStatisticsDTO.setPay_ratio( (new BigDecimal(customerStatisticsDTO.getOrder_amount()).divide(new BigDecimal(customerStatisticsDTO.getMember_count()),6,RoundingMode.HALF_UP)).doubleValue());



        SystemLogUtil.info("耗时：" + Duration.between(inst1, Instant.now()).getSeconds());
        List<CustomerStatisticsDTO> customerStatisticsDTOList=new ArrayList<>();
        customerStatisticsDTOList.add(customerStatisticsDTO);

        webPager.setParam(1,1,1,customerStatisticsDTOList);

        return Result.buildGrid(webPager);
    }
   
    
    
