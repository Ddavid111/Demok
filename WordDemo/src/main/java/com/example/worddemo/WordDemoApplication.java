package com.example.worddemo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import com.example.worddemo.Generate_Document;
@SpringBootApplication
public class WordDemoApplication {

    @Autowired
    private static Generate_Document generate_document;
    public static void main(String[] args) {
        SpringApplication.run(WordDemoApplication.class, args);
        generate_document.Proba();


    }




}
