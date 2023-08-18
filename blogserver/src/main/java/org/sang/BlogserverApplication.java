package org.sang;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;
@SpringBootApplication
@EnableScheduling//开启定时任务支持
public class BlogserverApplication
/*        extends SpringBootServletInitializer*/
{
 /*   @Override
    protected SpringApplicationBuilder configure(SpringApplicationBuilder builder) {
        // 注意这里要指向原先用main方法执行的Application启动类
        return builder.sources(BlogserverApplication.class);
    }*/
    public static void main(String[] args) {
        SpringApplication.run(BlogserverApplication.class, args);
    }
}





