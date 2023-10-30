package org.example;


import cn.hutool.core.thread.ConcurrencyTester;
import cn.hutool.core.thread.ThreadUtil;
import cn.hutool.core.util.IdUtil;


public class Main {
    public static void main(String[] args) throws Exception {
//        threadSize：并发数量
        ConcurrencyTester test = ThreadUtil.concurrencyTest(100, () -> {
//        并发业务代码
                    long snowflakeNextId = IdUtil.getSnowflakeNextId();
                    System.out.println(Thread.currentThread().getName() + "============" + snowflakeNextId);
                }
        );
//       统计并发解决响应时间
        long interval = test.getInterval();
        System.out.println("interval: " + interval);
//       关闭测试多线程线程池
        test.close();
    }
}