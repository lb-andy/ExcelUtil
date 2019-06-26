package me.qping.utils.excel.common;

import java.util.Collection;

/**
 * @ClassName Consumer
 * @Author qping
 * @Date 2019/5/15 03:48
 * @Version 1.0
 **/
public interface Consumer<T> {

    public void execute(Collection<T> data, Config config);
}
