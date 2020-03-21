/*
 * Copyright 2014-2020 Sayi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.deepoove.poi.util;

import com.deepoove.poi.exception.ReflectionException;
import org.apache.commons.lang3.ClassUtils;

import java.lang.reflect.Field;
import java.util.Objects;

public class ReflectionUtils {

    public static Object getValue(String fieldName, Object obj) {
        Objects.requireNonNull(obj, "Class must not be null");
        Objects.requireNonNull(fieldName, "Name must not be null");
        Field field = findField(obj.getClass(), fieldName);
        if (null == field) {
            throw new ReflectionException(
                    "No Such field " + fieldName + " from class" + ClassUtils.getShortClassName(obj.getClass()));
        }
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (Exception e) {
            throw new ReflectionException(fieldName, obj.getClass(), e);
        }
    }

    public static Field findField(Class<?> clazz, String name) {
        Objects.requireNonNull(clazz, "Class must not be null");
        Objects.requireNonNull(name, "Name must not be null");
        Class<?> searchType = clazz;
        while (Object.class != searchType && searchType != null) {
            Field field;
            try {
                field = searchType.getDeclaredField(name);
                if (null != field) return field;
            } catch (NoSuchFieldException e) {
                // no-op
            }
            searchType = searchType.getSuperclass();
        }
        return null;
    }

}
