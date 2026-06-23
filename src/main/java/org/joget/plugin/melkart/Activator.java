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

        // Register plugin classes here
        registrations.add(context.registerService(ExcelParser.class.getName(), new ExcelParser(), null));
    }

    public void stop(BundleContext context) {
        for (ServiceRegistration registration : registrations) {
            registration.unregister();
        }
    }
}
