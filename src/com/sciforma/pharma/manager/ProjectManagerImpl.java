/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package com.sciforma.pharma.manager;


import com.sciforma.pharma.exeception.TechnicalException;
import com.sciforma.psnext.api.AccessException;
import com.sciforma.psnext.api.InvalidDataException;
import com.sciforma.psnext.api.PSException;
import com.sciforma.psnext.api.Project;
import com.sciforma.psnext.api.Session;
//import fr.sciforma.psconnect.service.exception.BusinessException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 *
 * @author lahoudie
 */
public class ProjectManagerImpl implements ProjectManager {

	private Map<String, Project> projectsCacheMap = new HashMap<String, Project>();

	private Project nonProject;

	private boolean hasProjectCache = true;

	private boolean useSkeleton = false;

	private int version = Project.VERSION_WORKING;

	private int access = Project.READWRITE_ACCESS;

	private Session session;

	public ProjectManagerImpl(Session session) {
		this.session = session;
	}

	/**
	 * @param version
	 *            the version to set
	 */
	public void setVersion(int version) {
		this.version = version;
	}

	/**
	 * fluent setter
	 *
	 * @param versionPublished
	 * @return this
	 */
	public ProjectManagerImpl withVersion(int version) {
		setVersion(version);
		return this;
	}

	/**
	 * @param access
	 *            the access to set
	 */
	public void setAccess(int access) {
		this.access = access;
	}

	/**
	 * fluent setter access
	 *
	 * @param access
	 *            the access to set
	 */
	public ProjectManagerImpl withAccess(int access) {
		this.access = access;
		return this;
	}

	/**
	 * @param useSkeleton
	 *            the useSkeleton to set
	 */
	public void setUseSkeleton(boolean useSkeleton) {
		this.useSkeleton = useSkeleton;
	}

	/**
	 * fluent setter useSkeleton
	 *
	 * @param useSkeleton
	 * @return
	 */
	public ProjectManagerImpl withUseSkeleton(boolean useSkeleton) {
		this.useSkeleton = useSkeleton;
		return this;
	}

	/**
	 * @param hasProjectCache
	 *            the hasProjectCache to set
	 */
	public void setHasProjectCache(boolean hasProjectCache) {
		this.hasProjectCache = hasProjectCache;
	}

	/**
	 * fluent setter
	 *
	 * @param hasProjectCache
	 *            the hasProjectCache to set
	 * @return this
	 */
	public ProjectManagerImpl withHasProjectCache(boolean hasProjectCache) {
		this.hasProjectCache = hasProjectCache;
		return this;
	}

	private void putInCache(String id, Project lProject) {
		if (lProject.isNonProject()) {
			nonProject = lProject;
		} else {
			projectsCacheMap.put(id, lProject);
		}
	}

	public Project findProjectById(final String id) throws TechnicalException {
		if (hasProjectCache) {
			Project project = projectsCacheMap.get(id);

			if (project != null) {
				return project;
			}
		}

		List<Project> findAndCacheByFilter = findAndCacheByFilter(
				new ProjectFilter() {
					public boolean filter(Project project) throws PSException {
						return id.equals(project.getStringField("ID"));
					}
				}, true);

		return findAndCacheByFilter.isEmpty() ? null : findAndCacheByFilter
				.iterator().next();
	}

	public List<Project> findProjectByCriteria(ProjectFilter projectFilter)
			throws TechnicalException {
		return findAndCacheByFilter(projectFilter, false);
	}

	@SuppressWarnings("unchecked")
	private List<Project> findAndCacheByFilter(ProjectFilter projectFilter,
			boolean firstStop) {
		List<Project> list = new LinkedList<Project>();
		try {
			for (Project project : (List<Project>) (useSkeleton ? this.session
					.getSkeletonProjectList(version, access) : this.session
					.getProjectList(version, access))) {
				try {
					if (projectFilter.filter(project)) {
						list.add(project);
						if (firstStop) {
							return list;
						}
					}
					if (hasProjectCache) {
						putInCache(project.getStringField("ID"),
								project);
					}
				} catch (PSException e) {
					throw new TechnicalException(e);
				}
			}
			if (!projectFilter.onlyActived()) {
				for (Project project : (List<Project>) this.session
						.getInactiveProjectList(version, access)) {
					try {
						if (projectFilter.filter(project)) {
							list.add(project);
							if (firstStop) {
								return list;
							}
						}
						if (hasProjectCache) {
							putInCache(project.getStringField("ID"),
									project);
						}
					} catch (PSException e) {
						throw new TechnicalException(e);
					}
				}
			}
		} catch (AccessException e) {
			throw new TechnicalException(e,
					"problème d'accès lors de la recherche de projet <"
							+ projectFilter + ">");
		} catch (InvalidDataException e) {
			throw new TechnicalException(e,
					"problème de données invalide lors de la recherche de projet <"
							+ projectFilter + ">");
		} catch (PSException e) {
			throw new TechnicalException(e);
		}
		return list;
	}

}
