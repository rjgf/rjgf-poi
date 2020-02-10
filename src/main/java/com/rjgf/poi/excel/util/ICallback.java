package com.rjgf.poi.excel.util;

import java.util.List;

public interface ICallback<T> {

    void todo(List<T> data);

    Class getTClass();
}
