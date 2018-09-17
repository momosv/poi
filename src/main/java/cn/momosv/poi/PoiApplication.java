package cn.momosv.poi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;

@ConfigurationProperties("application.yml") //接收application.yml中的myProps下面的属性
@SpringBootApplication(exclude = DataSourceAutoConfiguration.class)
public class PoiApplication extends SpringBootServletInitializer {

    public static void main(String[] args) {
        SpringApplication.run(PoiApplication.class, args);
    }
}
