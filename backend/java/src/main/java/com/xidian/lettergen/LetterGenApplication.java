package com.xidian.lettergen;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * 制函服务 Java 后端（功能对齐 Python 版 backend/python/app.py，端口 5002）。
 * 与 Python 后端互斥运行；共用 frontend/ 与 templates_store/ 存储结构。
 */
@SpringBootApplication
public class LetterGenApplication {
    public static void main(String[] args) {
        SpringApplication.run(LetterGenApplication.class, args);
    }
}
