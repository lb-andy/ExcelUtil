package me.qping.utils.excel.anno;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    public String name() default "";
    public int index() default -1;
    public String dateformat() default "";
    public int sort() default -1;

}
