package com.zl.controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author zhuolin
 * @program: poi-demo
 * @date 2018/11/27
 * @description: ${description}
 **/
@RestController
@RequestMapping(value = "/api")
public class DemoController {

    @RequestMapping(value = "/word2007Html")
    public void word2007Html(){
    }
}
