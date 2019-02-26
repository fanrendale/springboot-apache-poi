package com.xjf.apachepoitest.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

/**
 * @author xjf
 * @date 2019/2/25 16:14
 */
@Controller
public class HelloController {


    @GetMapping("/hello")
    @ResponseBody
    public String hello(){
        return "Hello World";
    }
}
