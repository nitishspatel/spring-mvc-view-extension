package org.springframework.view.extension.config;

/**
 * Annotation to mark a field of <ReportBean> as a column in resultant csv/Excel file.
 * 
 * @author nitish
 *
 */
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Header {
	 /**
     * @return - The title of the column in csv.
     */
    String title() default "";
    
    /**
     * @return - The columnNumber in the csv.
     */
    int columnNumber() default 0;
    
    /**
     * @return - The columnNumber.
     */
    String colorFormatter() default "";
    String headerFormatter() default "";
    
    Class type() default String.class;
    
}
