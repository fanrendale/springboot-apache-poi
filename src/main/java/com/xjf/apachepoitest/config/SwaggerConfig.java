package com.xjf.apachepoitest.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import springfox.documentation.builders.ApiInfoBuilder;
import springfox.documentation.builders.PathSelectors;
import springfox.documentation.builders.RequestHandlerSelectors;
import springfox.documentation.service.ApiInfo;
import springfox.documentation.service.Contact;
import springfox.documentation.spi.DocumentationType;
import springfox.documentation.spring.web.plugins.Docket;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

/**
 * @author xjf
 * @date 2019/2/25 16:20
 */
@Configuration      //配置类
@EnableSwagger2     //启用Swagger2
public class SwaggerConfig {
    @Bean
    public Docket dataManagerApi(){
        return new Docket(DocumentationType.SWAGGER_2)
                .groupName("POI解析word文档")
                .apiInfo(apiInfo())
                .select()
                .apis(RequestHandlerSelectors.basePackage("com.xjf.apachepoitest.controller"))
                .paths(PathSelectors.any()).build();
    }

    private ApiInfo apiInfo(){
        return new ApiInfoBuilder()
                .title("POI解析word文档")
                .description("POI解析word文档，包括doc和docx格式")
                .termsOfServiceUrl("-----")
                .contact(new Contact("xjf","----","1053314919@qq.com"))
                .version("1.0").build();
    }
}
