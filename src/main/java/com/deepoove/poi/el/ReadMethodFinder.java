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
package com.deepoove.poi.el;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;

class ReadMethodFinder {

    private static Logger logger = LoggerFactory.getLogger(ReadMethodFinder.class);

    public static Method find(Class<?> objClass, String key) {
        try {
            PropertyDescriptor propDesc = new PropertyDescriptor(key, objClass);
            return propDesc.getReadMethod();
        } catch (IntrospectionException e1) {
            logger.debug("Fail introspector the property: {} from {}, {}", key, objClass, e1.getMessage());
        }
        return null;
    }

}
