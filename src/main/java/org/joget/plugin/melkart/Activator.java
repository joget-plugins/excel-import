package org.joget.plugin.melkart;

import java.util.ArrayList;
import java.util.Collection;
import org.osgi.framework.BundleActivator;
import org.osgi.framework.BundleContext;
import org.osgi.framework.ServiceRegistration;

public class Activator implements BundleActivator {

    protected Collection<ServiceRegistration> registrations;

    public void start(BundleContext context) {
        registrations = new ArrayList<ServiceRegistration>();

        // Register plugin classes here.
        // ExcelImportBinder is intentionally NOT registered: the element instantiates it directly
        // in getStoreBinder()/getLoadBinder(), so it never needs to be discovered by class name
        // (and we don't want it appearing as a selectable binder in the form builder).
        registrations.add(context.registerService(ExcelImport.class.getName(), new ExcelImport(), null));
    }

    public void stop(BundleContext context) {
        for (ServiceRegistration registration : registrations) {
            registration.unregister();
        }
    }
}
