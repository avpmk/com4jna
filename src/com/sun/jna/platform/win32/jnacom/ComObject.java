/*
 * Copyright 2009 Digital Rapids Corp.
 * 
 */

package com.sun.jna.platform.win32.jnacom;

import com.sun.jna.Function;
import com.sun.jna.Pointer;
import com.sun.jna.Structure;
import com.sun.jna.WString;
import com.sun.jna.platform.win32.Guid;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.Ole32Util;
import com.sun.jna.platform.win32.Oleaut32;
import com.sun.jna.platform.win32.W32API.HRESULT;
import com.sun.jna.ptr.PointerByReference;
import com.sun.jna.ptr.IntByReference;
import com.sun.jna.ptr.LongByReference;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author scott.palmer@digital-rapids.com
 */
public class ComObject implements InvocationHandler {
    static Ole32 OLE32 = Ole32.INSTANCE;

    static final int ptrSize = Pointer.SIZE;
    public static final int CLSCTX_INPROC_SERVER = 0x1;
    public static final int CLSCTX_INPROC_HANDLER = 0x2;
    public static final int CLSCTX_LOCAL_SERVER = 0x4;
    public static final int CLSCTX_INPROC_SERVER16 = 0x8;
    public static final int CLSCTX_REMOTE_SERVER = 0x10;
    public static final int CLSCTX_ALL = (CLSCTX_INPROC_SERVER
            | CLSCTX_INPROC_HANDLER
            | CLSCTX_LOCAL_SERVER
            | CLSCTX_REMOTE_SERVER);

    // TODO: use thread local storage for this
    private static int lastHRESULT = 0;
    public static final Guid.GUID IID_IUnknown = Ole32Util.getGUIDFromString("{00000000-0000-0000-C000-000000000046}");

    private Pointer _InterfacePtr = null;

    public ComObject(Pointer interfacePointer) {
        _InterfacePtr = interfacePointer;
    }

    public static<T extends IUnknown> T createInstance(Class<T> primaryInterface, String clsid) {
        Guid.GUID refclsid = Ole32Util.getGUIDFromString(clsid);
        Guid.GUID refiid = IID_IUnknown;
        try {
            String iid = (String) primaryInterface.getAnnotation(IID.class).value();
            refiid = Ole32Util.getGUIDFromString(iid);
        } catch (Exception ex) {
            Logger.getLogger(ComObject.class.getName()).log(Level.SEVERE, null, ex);
        }

        PointerByReference punkown = new PointerByReference();
        HRESULT hresult = OLE32.CoCreateInstance(refclsid, Pointer.NULL, CLSCTX_ALL, refiid, punkown);
        lastHRESULT = hresult.intValue();
        if (hresult.intValue() < 0)
            throw new ComException("CoCreateInstance returned 0x"+Integer.toHexString(hresult.intValue()),hresult.intValue());
        Pointer interfacePointer = punkown.getValue();
        //Function;
        return createProxyForNewObject(new ComObject(interfacePointer), primaryInterface);
    }

    public static int getLastHRESULT() {
        return lastHRESULT;
    }

    private <T> T createProxy(Class<?>... interfaces) {
        T p = (T) Proxy.newProxyInstance(ComObject.class.getClassLoader(), interfaces, this);
        return p;
    }

    private static<T> T createProxyForNewObject(ComObject object, Class<T> interfaces) {
        T p = (T) Proxy.newProxyInstance(ComObject.class.getClassLoader(), new Class<?>[] {interfaces}, object);
        return p;
    }

    /*
     * QueryInterface(REFIID, void **ppvObject)
     * AddRef(void)
     * Release(void)
     */
    private Pointer queryInterface(Class<?> comInterface) {
        try {
            String iid = (String) comInterface.getAnnotation(IID.class).value();
            Pointer vptr = _InterfacePtr.getPointer(0);
            Function func = Function.getFunction(vptr.getPointer(0));
            PointerByReference ppvObject = new PointerByReference();
            Guid.GUID refiid = Ole32Util.getGUIDFromString(iid);
            int hresult = func.invokeInt(new Object[]{_InterfacePtr, refiid, ppvObject});
            lastHRESULT = hresult;
            if (hresult >= 0) {
                return ppvObject.getValue();
            }
            throw new ComException("queryInterface failed. HRESULT = 0x"+Integer.toHexString(hresult),hresult);
        } catch (IllegalArgumentException ex) {
            Logger.getLogger(ComObject.class.getName()).log(Level.SEVERE, null, ex);
            throw new RuntimeException("queryInterface failed",ex);
        }
    }


    //
    // When preparing args, if a parameter is a COM interface replace it with
    // its _InterfacePtr pointer
    //
    //
    Object [] prepareArgs(Method method, Object [] args) {
        int asize = 1 + (args != null ? args.length : 0);
        Object[] aarg = new Object[asize];
        for (int i = 1; i < aarg.length; i++) {
            Object givenArg = args[i - 1];
            if (givenArg instanceof IUnknown) {
                aarg[i] = ((ComObject)Proxy.getInvocationHandler(givenArg))._InterfacePtr;
            } else {
                aarg[i] = givenArg;
            }
        }
        aarg[0] = _InterfacePtr;
        return aarg;
    }

    /**
     * Add the "this" pointer to the argument list and reserve a spot for the
     * "return" value.
     * @param method
     * @param args
     * @return the new arguments
     */
    Object[] prepareArgsPlusRetVal(Method method, Object[] args) {
        Object[] aarg;
        if (args != null) {
            // add two, one for the 'this' pointer, the other for the retVal
            aarg = new Object[2 + args.length];
            for (int i = 0; i < args.length; i++) {
                Object givenArg = args[i];
                if (givenArg instanceof IUnknown) {
                    aarg[i+1] = ((ComObject) Proxy.getInvocationHandler(givenArg))._InterfacePtr;
                } else {
                    aarg[i+1] = givenArg;
                }
            }
        } else {
            // Even though the method was declared with no arguments, we need
            // two.  One for the 'this' pointer, the other for the retVal.
            aarg = new Object[2];
        }
        // 'this' pointer is taked on to the "front", retVal will be last
        aarg[0] = _InterfacePtr;
        return aarg;
    }

    /**
     * Augment the parameter list for a COM API call that returns an integer
     * by adding the "this" pointer and the return value placeholder as a
     * byReference parameter.
     * @param method
     * @param args
     * @param retVal
     * @return
     */
    Object [] prepareArgs(Method method, Object [] args, IntByReference retVal) {
        Object[] aarg;
        ReturnValue rv = method.getAnnotation(ReturnValue.class);
        if (rv != null && rv.inout()) {
            aarg = prepareArgs(method, args);
            // replace the return value (in/out) reference
            retVal.setValue(((Number) args[rv.index()]).intValue());
            aarg[rv.index()] = retVal;
        } else {
            // default is to add the return value at the end of the parameter
            // list if it is not explicite
            aarg = prepareArgsPlusRetVal(method, args);
            aarg[aarg.length-1] = retVal;
        }
        aarg[0] = _InterfacePtr;
        return aarg;
    }

    Object[] prepareArgs(Method method, Object[] args, LongByReference retVal) {
        Object[] aarg;
        ReturnValue rv = method.getAnnotation(ReturnValue.class);
        if (rv != null && rv.inout()) {
            aarg = prepareArgs(method, args);
            // replace the return value (in/out) reference
            retVal.setValue(((Number) args[rv.index()]).longValue());
            aarg[rv.index()] = retVal;
        } else {
            aarg = prepareArgsPlusRetVal(method, args);
            aarg[aarg.length-1] = retVal;
        }
        aarg[0] = _InterfacePtr;
        return aarg;
    }

    Object[] prepareArgs(Method method, Object[] args, PointerByReference retVal) {
        Object[] aarg;
        ReturnValue rv = method.getAnnotation(ReturnValue.class);
        if (rv != null && rv.inout()) {
            aarg = prepareArgs(method, args);
            // replace the return value (in/out) reference
            retVal.setValue((Pointer) args[rv.index()]);
            aarg[rv.index()] = retVal;
        } else {
            aarg = prepareArgsPlusRetVal(method, args);
            aarg[aarg.length-1] = retVal;
        }
        aarg[0] = _InterfacePtr; // pass the "this" pointer
        return aarg;
    }
    
    Object[] prepareArgsObjOut(Method method, Object[] args, Object retVal) {
        Object[] aarg;
        ReturnValue rv = method.getAnnotation(ReturnValue.class);
        if (rv != null && rv.inout()) {
            aarg = prepareArgs(method, args);
            // replace the return value (in/out) reference
            //retVal.setValue((Pointer) args[rv.index()]);
            aarg[rv.index()] = retVal;
        } else {
            aarg = prepareArgsPlusRetVal(method, args);
            aarg[aarg.length-1] = retVal;
        }
        aarg[0] = _InterfacePtr; // pass the "this" pointer
        return aarg;
    }

    void invokeVoidCom(Method method, Object... args) {
        int offset = method.getAnnotation(VTID.class).value();
        Pointer vptr = _InterfacePtr.getPointer(0);
        Function func = Function.getFunction(vptr.getPointer(offset * ptrSize));
        Object[] aarg = prepareArgs(method, args);
        int hresult = func.invokeInt(aarg);
        lastHRESULT = hresult;
        if (hresult < 0)
            throw new ComException("Invocation of \"" + method.getName() + "\" failed, hresult=0x" + Integer.toHexString(hresult), hresult);
    }

    int invokeIntCom(Method method, Object... args){
        int offset = method.getAnnotation(VTID.class).value();
        Pointer vptr = _InterfacePtr.getPointer(0);
        Function func = Function.getFunction(vptr.getPointer(offset * ptrSize));
        IntByReference retVal = new com.sun.jna.ptr.IntByReference();
        Object[] aarg = prepareArgs(method, args, retVal);
        int hresult = func.invokeInt(aarg);
        lastHRESULT = hresult;
        if (hresult < 0)
            throw new ComException("Invocation of \"" + method.getName() + "\" failed, hresult=0x" + Integer.toHexString(hresult),hresult);
        return retVal.getValue();
    }

    long invokeLongCom(Method method, Object... args){
        int offset = method.getAnnotation(VTID.class).value();
        Pointer vptr = _InterfacePtr.getPointer(0);
        Function func = Function.getFunction(vptr.getPointer(offset * ptrSize));
        LongByReference retVal = new com.sun.jna.ptr.LongByReference();
        Object[] aarg = prepareArgs(method, args, retVal);
        int hresult = func.invokeInt(aarg);
        lastHRESULT = hresult;
        if (hresult < 0)
            throw new ComException("Invocation of \"" + method.getName() + "\" failed, hresult=0x" + Integer.toHexString(hresult), hresult);
        return retVal.getValue();
    }

    Object invokeObjectCom(Method method, Object... args) {
        int offset = method.getAnnotation(VTID.class).value();
        Pointer vptr = _InterfacePtr.getPointer(0);
        Function func = Function.getFunction(vptr.getPointer(offset * ptrSize));
        if (method.getReturnType().isInterface()) {
            // Hack the return parameter on to the args
            PointerByReference p = new PointerByReference();
            Object[] aarg = prepareArgs(method, args, p);
            int hresult = (Integer) func.invoke(Integer.class, aarg);
            lastHRESULT = hresult;
            if (hresult < 0)
                throw new ComException("Invocation of \""+method.getName()+"\" failed, hresult=0x"+Integer.toHexString(hresult), hresult);
            return createProxyForNewObject(new ComObject(p.getValue()), method.getReturnType());
        } else {
            boolean returnsBSTR = false;
            Object retVal = null;
            ReturnValue rv = method.getAnnotation(ReturnValue.class);
            if (rv != null && rv.inout()) {
                // one of the args is the return value
                retVal = (Structure) args[rv.index()];
            } else {
                // retval wasn't an explicit parameter - so we will make one
                if (method.getReturnType() == String.class) {
                    // String in COM should mean BSTR
                    returnsBSTR = true;
                    retVal = new PointerByReference();
                } else try {
                    // just a normal structure
                    retVal = method.getReturnType().newInstance();
                    assert retVal instanceof Structure;
                } catch (InstantiationException ex) {
                    Logger.getLogger(ComObject.class.getName()).log(Level.SEVERE, null, ex);
                    throw new RuntimeException("Invocation of \"" + method.getName() + "\" failed.",ex);
                } catch (IllegalAccessException ex) {
                    Logger.getLogger(ComObject.class.getName()).log(Level.SEVERE, null, ex);
                    throw new RuntimeException("Invocation of \"" + method.getName() + "\" failed.", ex);
                }
            }
            Object[] aarg = prepareArgsObjOut(method, args, retVal);
            int hresult = (Integer) func.invoke(Integer.class, aarg);
            lastHRESULT = hresult;
            if (hresult < 0)
                throw new ComException("Invocation of \"" + method.getName() + "\" failed, hresult=0x" + Integer.toHexString(hresult), hresult);
            if (returnsBSTR) {
                Pointer bstr = ((PointerByReference) retVal).getValue();
                retVal = bstr.getString(0,true);
                Oleaut32.INSTANCE.SysFreeString(bstr);
            }
            return retVal;
        }
    }

    void addRef() {
        Pointer vptr = _InterfacePtr.getPointer(0);
        Function func = Function.getFunction(vptr.getPointer(ptrSize));
        func.invoke(new Object[] {_InterfacePtr});
    }
    
    void release() {
        Pointer vptr = _InterfacePtr.getPointer(0);
        Function func = Function.getFunction(vptr.getPointer(2 * ptrSize));
        func.invoke(new Object[]{_InterfacePtr});
    }

    /**
     * The "Java" way to free resources is to call a dispose() method, after
     * which, the object should not be used.  Since Java passes only object
     * references by value, the need for explicite addRef and release calls
     * is rare.
     */
    public void dispose() {
        if (_InterfacePtr != null) {
            release();
            _InterfacePtr = null;
        }
    }

    @Override
    protected void finalize() throws Throwable {
        dispose();
        super.finalize();
    }

    public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
        if (method.getName().equals("queryInterface")) {
            // if the native interface pointer is the same, return a new proxy to
            // this same ComObject that implements the interface required interface
            // otherwise make a new ComObject to wrap the returned interface pointer
            if (args[0] instanceof Class<?>) {
                Class<?> c = (Class<?>) args[0];
                if (c.isInterface()) {
                    Pointer interfacePointer = queryInterface((Class<?>)args[0]);
                    return createProxyForNewObject(new ComObject(interfacePointer), c);
                } else {
                    throw new RuntimeException("Argument to queryInterface must be a Java interface class annotated with an interface ID.");
                }
            }
        } else if (method.getName().equals("dispose")) {
            //System.out.println("dispose() ->"+proxy);
            dispose();
            return null;
        } else if (method.getName().equals("toString")) {
            StringBuilder sb = new StringBuilder();
            sb.append(_InterfacePtr);
            sb.append("):");

            boolean notFirst = false;
            for (Class<?> cc : proxy.getClass().getInterfaces()) {
                if (notFirst)
                    sb.append(", ");
                else
                    notFirst = true;
                sb.append(cc.getName());
            }

            return sb.toString();
        }
       
        //Examine the args and wrap String objects as WString
        if(args != null) {
            for(int i=0;i<args.length;i++) {
                if(args[i] instanceof String) {
                    String s = (String)(args[i]);
                    WString wstr = new WString(s);
                    args[i] = wstr;
                }
            }
        }

        if (method.getReturnType() == Void.TYPE) {
            invokeVoidCom(method, args);
            return null;
        } else if (method.getReturnType() == Integer.TYPE) {
            return invokeIntCom(method, args);
        } else if (method.getReturnType() == Long.TYPE) {
            return invokeLongCom(method, args);
        } else {
            return invokeObjectCom(method, args);
        }
    }
}