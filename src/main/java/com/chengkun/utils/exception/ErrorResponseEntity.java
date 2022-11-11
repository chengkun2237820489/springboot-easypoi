package com.chengkun.utils.exception;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 异常信息实体类
 *
 * @author chengkun
 * @version v1.0
 * @create 2020/8/1 23:34
 **/
@Data
@AllArgsConstructor
public class ErrorResponseEntity {
    private int code;
    private String message;
}
