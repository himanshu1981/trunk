package com.example.excelexport.config;

import org.springframework.context.annotation.Configuration;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;
import org.springframework.web.servlet.config.annotation.AsyncSupportConfigurer;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

/**
 * Configures the async task executor used by Spring MVC's StreamingResponseBody.
 *
 * Without this, large exports run on the container's main request thread pool,
 * which can exhaust connections while waiting for a slow query/write to finish.
 * A dedicated small pool (2-5 threads) isolates that back-pressure.
 *
 * Timeout is set to -1 (infinite) so slow exports of 1 M+ rows are not cut off.
 */
@Configuration
public class ExcelAsyncConfig implements WebMvcConfigurer {

    @Override
    public void configureAsyncSupport(AsyncSupportConfigurer configurer) {
        configurer.setDefaultTimeout(-1); // no timeout — export can take minutes

        ThreadPoolTaskExecutor executor = new ThreadPoolTaskExecutor();
        executor.setCorePoolSize(2);
        executor.setMaxPoolSize(5);
        executor.setQueueCapacity(10);
        executor.setThreadNamePrefix("excel-export-");
        executor.initialize();
        configurer.setTaskExecutor(executor);
    }
}
