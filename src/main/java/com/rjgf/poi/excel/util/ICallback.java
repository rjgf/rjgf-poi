package com.rjgf.poi.excel.util;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.List;

public interface ICallback<T> {

    void todo(List<T> data);

    default Class getTClass() {
        Type[] types = this.getClass().getGenericInterfaces();
        System.out.println("["+types[0].getTypeName()+"]");
        if (!(types[0] instanceof ParameterizedType)) {
            throw new RuntimeException("接口不是泛型类 ["+types[0].getTypeName()+"]");
        }
        ParameterizedType parameterized = (ParameterizedType) types[0];
        return (Class<T>) parameterized.getActualTypeArguments()[0];
    }
}
